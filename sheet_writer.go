package xlsx_tags

// provides helper utility for writing to the xls sheet.
// Available tag options:
// 	order - positive number, defines cells ordering, required,
// 	heading - column title, required
// 	format - optional formatting

import (
	"errors"
	"fmt"
	"github.com/tealeg/xlsx"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"
)

const tagName = "xls"

var (
	ErrUnsupportedType        = errors.New("trying to marshal unsupported type. Supported types are: array, slice, struct")
	ErrUnsupportedContentType = errors.New("trying to marshal unsupported content type. Currently supports only struct")
	ErrHeadingPropRequired    = errors.New("heading property must be set")
)

var marshallerType = reflect.TypeOf((Marshaller)(nil))

// todo think about name correctness
type parseOpts struct {
	order   int
	heading string
	format  string
}

func (o parseOpts) hasFormatting() bool {
	return o.format == ""
}

type optsByOrder []parseOpts

func (o optsByOrder) Len() int           { return len(o) }
func (o optsByOrder) Swap(i, j int)      { o[i], o[j] = o[j], o[i] }
func (o optsByOrder) Less(i, j int) bool { return o[i].order < o[j].order }

func WriteToSheet(sheet *xlsx.Sheet, data interface{}) error {
	//fixme avoid code duplication
	v := reflect.ValueOf(data)
	kind := v.Kind()
	if kind != reflect.Array && kind != reflect.Slice {
		return ErrUnsupportedType
	}

	itemsType := getListType(data)
	if itemsType.Kind() != reflect.Struct {
		return ErrUnsupportedContentType
	}

	if v.Type().Implements(marshallerType) {
		return writeWithMarshaller(sheet, data)
	} else {
		return writeWithTags(sheet, data)
	}
}

func writeWithMarshaller(sheet *xlsx.Sheet, data interface{}) error {
	m := data.(Marshaller)
	writeHeading(sheet, m.Header())
	for _, row := range m.Data() {
		writeRow(sheet, row)
	}
	return nil
}

func writeWithTags(sheet *xlsx.Sheet, data interface{}) error {
	v := reflect.ValueOf(data)
	itemsType := getListType(data)
	opts, err := getParseOptions(itemsType)
	if err != nil {
		return err
	}
	sort.Sort(optsByOrder(opts))
	// stores mapping between order and cell position
	orderToCellPos := make(map[int]int)
	for i, opt := range opts {
		orderToCellPos[opt.order] = i
	}
	// write heading
	writeHeadingFromOpts(sheet, opts)

	values := make([]string, len(opts))
	for i := 0; i < v.Len(); i++ {
		item := v.Index(i)
		values = values[:]
		for j := 0; j < item.NumField(); j++ {
			tag := item.Type().Field(j).Tag.Get(tagName)

			// Skip if tag is not defined or ignored
			if tag == "" || tag == "-" {
				continue
			}

			// it will not trigger an error as we already verified it
			order, _ := orderFromTag(tag)
			pos := orderToCellPos[order]
			opt := opts[pos]
			switch val := item.Field(j).Interface().(type) {
			// todo maybe pass it to the default case
			case int, int8, int16, int32, int64, float32, float64, string:
				format := "%v"
				if opt.hasFormatting() {
					format = opt.format
				}
				values[pos] = fmt.Sprintf(format, val)
			case time.Time:
				var s string
				if opt.hasFormatting() {
					s = val.Format(opt.format)
				} else {
					s = val.String()
				}
				if val.IsZero() {
					s = " - "
				}
				values[pos] = s
			case fmt.Stringer:
				values[pos] = val.String()
			default:
				values[pos] = fmt.Sprintf("%v", val)
			}
		}
		writeRow(sheet, values)
	}
	return nil
}

func getParseOptions(data reflect.Type) ([]parseOpts, error) {
	var opts []parseOpts
	for i := 0; i < data.NumField(); i++ {
		tag := data.Field(i).Tag.Get(tagName)

		// Skip if tag is not defined or ignored
		if tag == "" || tag == "-" {
			continue
		}

		opt, err := parseOptFromTag(tag)
		if err != nil {
			return nil, err
		}
		opts = append(opts, opt)
	}
	return opts, nil
}

func writeHeadingFromOpts(sheet *xlsx.Sheet, opts []parseOpts) {
	var data = make([]string, len(opts))
	for key, opt := range opts {
		data[key] = opt.heading
	}
	writeRow(sheet, data)
}

func writeHeading(sheet *xlsx.Sheet, data []string) {
	writeRow(sheet, data)
}

func writeRow(sheet *xlsx.Sheet, data []string) {
	var cell *xlsx.Cell
	row := sheet.AddRow()
	for _, v := range data {
		cell = row.AddCell()
		cell.Value = v
	}
}

// todo modify options parsing
var (
	orderRegex   = regexp.MustCompile(`order=(?P<order>[\d\s\w]+),?|$`)
	headingRegex = regexp.MustCompile(`heading=(?P<heading>("[#\w.,-/\\ ]+")|([#\w\s]+))(,|$)`)
	formatRegex  = regexp.MustCompile(`format=(?P<format>[-\w\s%.\d\\/]+)(,|$)`)
)

func parseOptFromTag(tag string) (parseOpts, error) {
	order, err := orderFromTag(tag)
	if err != nil {
		return parseOpts{}, err
	}
	heading, err := headingFromTag(tag)
	if err != nil {
		return parseOpts{}, err
	}
	return parseOpts{
		order:   order,
		heading: heading,
		format:  formatFromTag(tag),
	}, nil
}

func orderFromTag(tag string) (int, error) {
	submatches := findStringSubmatchMap(orderRegex, tag)
	return strconv.Atoi(submatches["order"])
}

func headingFromTag(tag string) (string, error) {
	submatches := findStringSubmatchMap(headingRegex, tag)
	heading := strings.Trim(submatches["heading"], "\"")
	heading = strings.TrimSpace(heading)
	if heading != "" {
		return heading, nil
	}
	return heading, ErrHeadingPropRequired
}

func formatFromTag(tag string) string {
	submatches := findStringSubmatchMap(formatRegex, tag)
	return submatches["format"]
}

func findStringSubmatchMap(r *regexp.Regexp, s string) map[string]string {
	captures := make(map[string]string)

	match := r.FindStringSubmatch(s)
	if match == nil {
		return captures
	}

	for i, name := range r.SubexpNames() {
		// Ignore the whole regexp match and unnamed groups
		if i == 0 || name == "" {
			continue
		}
		captures[name] = match[i]
	}
	return captures
}

func getListType(s interface{}) reflect.Type {
	return reflect.TypeOf(s).Elem()
}
