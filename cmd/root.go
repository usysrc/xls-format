package cmd

import (
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "xls-format",
	Short: "Format Excel file table columns",
	Long: `A longer description that spans multiple lines and likely contains
examples and usage of using your application. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	// Run: func(cmd *cobra.Command, args []string) { },
	RunE: formatColumns,
}

// Declare the variables used in the format command
var (
	sheetIndex   int
	columnStart  string
	columnEnd    string
	columnFormat string
)

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	// Here you will define your flags and configuration settings.
	// Cobra supports persistent flags, which, if defined here,
	// will be global for your application.

	// rootCmd.PersistentFlags().StringVar(&cfgFile, "config", "", "config file (default is $HOME/.xls-format.yaml)")

	// Cobra also supports local flags, which will only run
	// when this action is called directly.
	rootCmd.SetUsageTemplate(`Usage: xls-format [file-path] [flags]
Flags:
  -s, --sheet int         Index of the sheet (starting from 0) (required)
  -b, --start string      Starting column (e.g., A) (required)
  -e, --end string        Ending column (e.g., Z) (required)
  -t, --format string     Column format: text or number (default "text")
`)
	rootCmd.Args = cobra.ExactArgs(1) // Expect exactly one argument
	rootCmd.Flags().IntVarP(&sheetIndex, "sheet", "s", 0, "Index of the sheet (starting from 0)")
	rootCmd.Flags().StringVarP(&columnStart, "start", "b", "", "Starting column (e.g., A)")
	rootCmd.Flags().StringVarP(&columnEnd, "end", "e", "", "Ending column (e.g., Z)")
	rootCmd.Flags().StringVarP(&columnFormat, "format", "t", "text", "Column format: text, number, or date")

	err := rootCmd.MarkFlagRequired("start")
	if err != nil {
		log.Fatal(err)
	}
	err = rootCmd.MarkFlagRequired("end")
	if err != nil {
		log.Fatal(err)
	}

}

func formatColumns(cmd *cobra.Command, args []string) error {
	filePath := args[0]
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}

	sheetName := f.GetSheetName(sheetIndex)

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}

	maxRow := len(rows)

	startColIndex, err := excelize.ColumnNameToNumber(columnStart)
	if err != nil {
		return fmt.Errorf("invalid start column: %s", columnStart)
	}

	endColIndex, err := excelize.ColumnNameToNumber(columnEnd)
	if err != nil {
		return fmt.Errorf("invalid end column: %s", columnEnd)
	}

	format, err := getColumnFormat(columnFormat)
	if err != nil {
		return err
	}

	style, err := f.NewStyle(&excelize.Style{
		NumFmt: format,
	})
	if err != nil {
		return err
	}

	for row := 1; row <= maxRow; row++ {
		for col := startColIndex; col <= endColIndex; col++ {
			cell := columnNumberToName(col) + strconv.Itoa(row)
			if err := f.SetCellStyle(sheetName, cell, cell, style); err != nil {
				return err
			}
		}
	}

	if err := f.Save(); err != nil {
		return err
	}

	fmt.Println("Formatting completed successfully.")
	return nil
}

func getColumnFormat(columnFormat string) (int, error) {
	formats := map[string]int{
		"text":   1,
		"number": 2,
		"date":   22,
	}
	format, ok := formats[columnFormat]
	if !ok {
		return 0, fmt.Errorf("unsupported column format: %v", columnFormat)
	}
	return format, nil
}

func columnNumberToName(col int) string {
	div := col
	name := ""
	for div > 0 {
		mod := (div - 1) % 26
		name = string(rune('A'+mod)) + name
		div = (div - mod) / 26
	}
	return name
}
