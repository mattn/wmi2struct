package main

import (
	"bytes"
	"encoding/xml"
	"flag"
	"fmt"
	"go/format"
	"io"
	"log"
	"os"
	"os/exec"
)

type Property struct {
	Name string `xml:"NAME,attr"`
	Type string `xml:"TYPE,attr"`
}

type class struct {
	Name     string     `xml:"NAME,attr"`
	Property []Property `xml:"PROPERTY"`
}

type result struct {
	Class class `xml:"RESULTS>CIM>CLASS"`
}

func typeName(s string) string {
	switch s {
	case "boolean":
		return "bool"
	case "char16":
		return "uint16"
	case "datetime":
		return "time.Time"
	case "object":
		return "interface{}"
	case "real32":
		return "float32"
	case "real64":
		return "float64"
	case "reference":
		return "interface{}"
	case "sint16":
		return "int16"
	case "sint32":
		return "int32"
	case "sint8":
		return "int8"
	case "string":
		return "string"
	case "uint8":
		return "uint8"
	case "uint16":
		return "uint16"
	case "uint32":
		return "uint32"
	case "uint64":
		return "uint16"
	}
	return s
}

func main() {
	var pkg, out string
	flag.StringVar(&pkg, "p", "main", "package")
	flag.StringVar(&out, "o", "", "output filename")
	flag.Parse()

	var buf bytes.Buffer

	fmt.Fprintf(&buf, "package %s\n", pkg)
	for _, arg := range flag.Args() {
		cmd := exec.Command("wmic", "class", arg, "get", "/format:RAWXML")
		b, err := cmd.CombinedOutput()
		if err != nil {
			log.Fatal(err)
		}
		var r result
		err = xml.Unmarshal(b, &r)
		if err != nil {
			log.Fatal(err)
		}
		fmt.Fprintf(&buf, "type %s struct {\n", r.Class.Name)
		for _, p := range r.Class.Property {
			fmt.Fprintf(&buf, "\t%s\t%s\n", p.Name, typeName(p.Type))
		}
		fmt.Fprintln(&buf, "}\n")
	}

	b, err := format.Source(buf.Bytes())
	if err != nil {
		log.Fatal(err)
	}

	var w io.Writer = os.Stdout
	if out != "" {
		f, err := os.Create(out)
		if err != nil {
			log.Fatal(err)
		}
		defer f.Close()
		w = f
	}
	w.Write(b)
}
