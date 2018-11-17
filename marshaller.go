package xlsx_tags

type Marshaller interface {
	Header() []string
	Data() [][]string
}
