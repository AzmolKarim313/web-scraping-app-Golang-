package main

import (
	"fmt"
	"log"

	"reflect"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gocolly/colly"
	"golang.org/x/exp/maps"
)

type product struct {
	ProductName string
	Rating      string
	Price       string
	// LinkURL     string
	Property map[string]string
}

const DOMAIN_NAME = "https://www.amazon.com"

func main() {
	// Colly’s main entity is a Collector object

	products := make([]product, 0)
	var product product

	c := colly.NewCollector()

	c.SetRequestTimeout(200 * time.Second)

	// Called before a request
	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL)
	})

	// Called if error occured during the request
	c.OnError(func(r *colly.Response, err error) {
		log.Println("========== ERROR START ==========")
		log.Printf("Something went wrong: \nCode : %v\nMessagr : %v\n", r.StatusCode, err)
		log.Println("========== ERROR END ==========")

		reVisit := r.Request.URL.String()
		fmt.Println("##### Re-Visiting URL : ", reVisit)
		c.AllowURLRevisit = true
		c.Visit(reVisit)
	})

	// Called after response received
	c.OnResponse(func(r *colly.Response) {
		fmt.Println("Visited", r.Request.URL)
	})

	c.OnHTML("table.a-keyvalue.prodDetTable", func(m *colly.HTMLElement) {
		var datas = map[string]string{}

		m.ForEach("tr", func(i int, h *colly.HTMLElement) {

			column := h.ChildText("th.a-color-secondary.a-size-base.prodDetSectionEntry")

			value := h.ChildText("td.a-size-base.prodDetAttrValue")

			if len(value) > 0 {
				datas[column] = value
			}
		})
		if len(datas) > 0 {
			product.Property = datas
		}
	})

	// Called right after OnResponse if the received content is HTML
	c.OnHTML("div.s-main-slot.s-result-list.s-search-results.sg-row", func(e *colly.HTMLElement) {
		e.ForEach("div.a-section.a-spacing-small.a-spacing-top-small", func(_ int, h *colly.HTMLElement) {
			name := h.ChildText("span.a-size-medium.a-color-base.a-text-normal")
			stars := h.ChildText("span.a-icon-alt")
			price := h.ChildText("span.a-price-whole") + h.ChildText("span.a-price-fraction")
			linkURL := h.ChildAttr("a.a-link-normal.s-underline-text.s-underline-link-text.s-link-style.a-text-normal", "href")

			if name != "" {
				product.ProductName = name
				product.Rating = stars
				product.Price = price

				h.Request.Visit(DOMAIN_NAME + linkURL)
			}
		})
	})

	c.OnHTML("a.s-pagination-next", func(e *colly.HTMLElement) {
		nextPage := e.Request.AbsoluteURL(e.Attr("href"))
		fmt.Printf("Next Page : %v\n", nextPage)
		c.Visit(nextPage)
	})

	// Called after OnHTML =>OnXML callbacks
	c.OnScraped(func(r *colly.Response) {
		products = append(products, product)
		c.AllowURLRevisit = false
	})

	// Start scraping on DOMAIN_NAME
	inputData(c)

	columns, err := getColumnNames(products)

	if err != nil {
		return
	}

	// fmt.Printf("\nAll Products are : %v\n", products)

	// fmt.Printf("\nColumns are : %v\n", columns)

	// writeExcelFile(columns, products)

	if len(products) > 1 {
		writeExcelFile(columns, unique(products))
	}

}

func inputData(c *colly.Collector) {
	f, err := excelize.OpenFile("InputData.xlsx")
	if err != nil {
		log.Fatalln(err)
	}
	firstSheet := f.WorkBook.Sheets.Sheet[0].Name

	rows := f.GetRows(firstSheet)

	for _, row := range rows {
		for _, colCell := range row {
			c.Visit(colCell)
		}
	}
}

func getColumnNames(datas []product) ([]string, error) {
	columnNames := []string{
		"Product Name",
		"Rating",
		"Price",
		// "Link URL",
	}

	for _, data := range datas {
		keys := maps.Keys(data.Property)

		for _, keyValue := range keys {
			if !Contains(columnNames, keyValue) {
				columnNames = append(columnNames, keyValue)
			}
		}
	}

	return columnNames, nil
}

func Contains[T comparable](s []T, e T) bool {
	for _, v := range s {
		if v == e {
			return true
		}
	}
	return false
}

func writeExcelFile(columns []string, datas []product) {
	file := excelize.NewFile()

	headerRowID := 1
	dataRowID := 2

	for i, column := range columns {
		file.SetCellValue("Sheet1", excelize.ToAlphaString(i)+strconv.Itoa(headerRowID), column)
		// fmt.Printf("Heading : %v and Column : %v\n", excelize.ToAlphaString(i)+strconv.Itoa(headerRowID), column)

	}

	for _, data := range datas {
		for i, column := range columns {
			if column == "Product Name" {
				file.SetCellValue("Sheet1", excelize.ToAlphaString(i)+strconv.Itoa(dataRowID), data.ProductName)
			} else if column == "Rating" {
				file.SetCellValue("Sheet1", excelize.ToAlphaString(i)+strconv.Itoa(dataRowID), data.Rating)
			} else if column == "Price" {
				file.SetCellValue("Sheet1", excelize.ToAlphaString(i)+strconv.Itoa(dataRowID), data.Price)
			} else {
				if len(data.Property) > 0 {
					keys := maps.Keys(data.Property)

					for _, keyValue := range keys {
						if column == keyValue {
							file.SetCellValue("Sheet1", excelize.ToAlphaString(i)+strconv.Itoa(dataRowID), data.Property[keyValue])
							break
						} else {
							continue
						}
					}
				} else {
					break
				}
			}
		}
		dataRowID++
	}

	if err := file.SaveAs("OutputData.xlsx"); err != nil {
		log.Fatal(err)
	}
}

func unique(s []product) []product {
	var results []product
	for _, product := range s {
		check := false

		for _, unique := range results {
			if reflect.DeepEqual(unique, product) {
				check = true
				break
			}
		}

		if !check {
			results = append(results, product)
		}
	}
	return results
}
