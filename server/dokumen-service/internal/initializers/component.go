package initializers

func GetStringOrNil(value string) *string {
	if value == "" {
		return nil
	}
	return &value
}

func GetColumn(row []string, index int) string {
	if index >= len(row) {
		return ""
	}
	return row[index]
}
