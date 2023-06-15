package cmd

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestColumnNumberToName(t *testing.T) {
	tests := []struct {
		columnNumber int
		expectedName string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
		{28, "AB"},
		{52, "AZ"},
		{53, "BA"},
		{702, "ZZ"},
		{703, "AAA"},
	}

	for _, tt := range tests {
		actualName := columnNumberToName(tt.columnNumber)
		assert.Equal(t, tt.expectedName, actualName, "Incorrect column name for column number %d", tt.columnNumber)
	}
}
