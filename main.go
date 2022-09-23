package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/measurement"
	"github.com/unidoc/unioffice/schema/soo/wml"
)

func init() {
	// Make sure to load your metered License API key prior to using the library.
	// If you need a key, you can sign up and create a free one at https://cloud.unidoc.io
	content, err := os.ReadFile("/home/yscrexm/Desktop/projects/generateOfficeFiles/apiKey")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(string(content))

	err = license.SetMeteredKey(string(content))
	if err != nil {
		panic(err)
	}

}

func schedule() {
	doc := document.New()
	defer doc.Close()

	//schedule
	{
		table := doc.AddTable()
		table.Properties().SetWidthPercent(100)
		borders := table.Properties().Borders()
		borders.SetAll(wml.ST_BorderSingle, color.Auto, 1*measurement.Point)

		row := table.AddRow()
		cell := row.AddCell()
		cell.Properties().SetWidthPercent(100)
		cell.AddParagraph().AddRun().AddText("schedule for " + time.Now().Format("02-01-2006"))

		//start new table
		doc.AddParagraph()
		table = doc.AddTable()
		table.Properties().SetWidthPercent(100)
		borders = table.Properties().Borders()
		borders.SetAll(wml.ST_BorderSingle, color.Auto, 1*measurement.Point)

		row = table.AddRow()
		cell = row.AddCell()
		cell.Properties().SetWidth(0.10 * measurement.Inch)
		cell.AddParagraph().AddRun().AddText("group#")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("0")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("1")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("2")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("3")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("4")

		row = table.AddRow()
		cell = row.AddCell()
		cell.Properties().SetWidth(0.10 * measurement.Inch)
		cell.AddParagraph().AddRun().AddText("55")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("math")
		cell.AddParagraph().AddRun().AddText("Mike")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("ukr")
		cell.AddParagraph().AddRun().AddText("Leroy")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("PE")
		cell.AddParagraph().AddRun().AddText("Jack")
		cell = row.AddCell()
		cell.AddParagraph().AddRun().AddText("")

	}
	doc.SaveToFile("docx/schedule.docx")
}

func main() {
	schedule()
}

// func createParaRun(doc *document.Document, s string) document.Run {
// 	para := doc.AddParagraph()
// 	run := para.AddRun()
// 	run.AddText(s)
// 	return run
// }
