//! encoding/* packages — one compile smoke per distinct API (breadth over depth).

use crate::helpers::*;

go_compile_cases! {
    xml_marshal => "package main; import \"encoding/xml\"; type T struct { X int `xml:\"x\"` }; func main() { _, _ = xml.Marshal(T{X: 1}) }",
    xml_unmarshal => "package main; import \"encoding/xml\"; type T struct { X int `xml:\"x\"` }; func main() { var t T; _ = xml.Unmarshal([]byte(`<T x=\"1\"/>`), &t) }",
    csv_new_reader => "package main; import \"encoding/csv\"; import \"strings\"; func main() { _ = csv.NewReader(strings.NewReader(\"a,b\")) }",
    csv_new_writer => "package main; import \"encoding/csv\"; import \"bytes\"; func main() { _ = csv.NewWriter(bytes.NewBuffer(nil)) }",
    gob_new_encoder => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { _ = gob.NewEncoder(bytes.NewBuffer(nil)) }",
    gob_new_decoder => "package main; import \"encoding/gob\"; import \"bytes\"; func main() { _ = gob.NewDecoder(bytes.NewBuffer(nil)) }",
    pem_encode => "package main; import \"encoding/pem\"; func main() { _ = pem.EncodeToMemory(&pem.Block{Type: \"TEST\", Bytes: []byte(\"x\")}) }",
    pem_decode => "package main; import \"encoding/pem\"; func main() { _, _ = pem.Decode([]byte(\"-----BEGIN TEST-----\\nx\\n-----END TEST-----\")) }",
    ascii85_encode => "package main; import \"encoding/ascii85\"; import \"bytes\"; func main() { _ = ascii85.NewEncoder(bytes.NewBuffer(nil)) }",
    base32_encode => "package main; import \"encoding/base32\"; func main() { _ = base32.StdEncoding.EncodeToString([]byte(\"go\")) }",
}
