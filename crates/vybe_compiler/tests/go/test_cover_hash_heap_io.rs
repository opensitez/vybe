//! hash/*, container/heap, compress/lzw, regexp/syntax — one API per test.

use crate::helpers::*;

go_compile_cases! {
    adler32_new => "package main; import \"hash/adler32\"; func main() { _ = adler32.New() }",
    adler32_checksum => "package main; import \"hash/adler32\"; func main() { _ = adler32.Checksum([]byte(\"go\")) }",
    crc64_new => "package main; import \"hash/crc64\"; func main() { _ = crc64.New(crc64.MakeTable(crc64.ISO)) }",
    crc64_checksum => "package main; import \"hash/crc64\"; func main() { _ = crc64.Checksum([]byte(\"go\"), crc64.MakeTable(crc64.ISO)) }",
    fnv_new32 => "package main; import \"hash/fnv\"; func main() { _ = fnv.New32() }",
    fnv_new32a => "package main; import \"hash/fnv\"; func main() { _ = fnv.New32a() }",
    fnv_new64 => "package main; import \"hash/fnv\"; func main() { _ = fnv.New64() }",
    fnv_new64a => "package main; import \"hash/fnv\"; func main() { _ = fnv.New64a() }",
    fnv_new128 => "package main; import \"hash/fnv\"; func main() { _ = fnv.New128() }",
    fnv_new128a => "package main; import \"hash/fnv\"; func main() { _ = fnv.New128a() }",
    heap_init => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i,j int) bool { return h[i]<h[j] }; func (h IH) Swap(i,j int) { h[i],h[j]=h[j],h[i] }; func (h *IH) Push(x interface{}) { *h=append(*h,x.(int)) }; func (h *IH) Pop() interface{} { o:=*h; n:=len(o); x:=o[n-1]; *h=o[:n-1]; return x }; func main() { h:=&IH{}; heap.Init(h) }",
    heap_push_pop => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i,j int) bool { return h[i]<h[j] }; func (h IH) Swap(i,j int) { h[i],h[j]=h[j],h[i] }; func (h *IH) Push(x interface{}) { *h=append(*h,x.(int)) }; func (h *IH) Pop() interface{} { o:=*h; n:=len(o); x:=o[n-1]; *h=o[:n-1]; return x }; func main() { h:=&IH{1}; heap.Init(h); heap.Push(h,2); _=heap.Pop(h) }",
    heap_fix_remove => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i,j int) bool { return h[i]<h[j] }; func (h IH) Swap(i,j int) { h[i],h[j]=h[j],h[i] }; func (h *IH) Push(x interface{}) { *h=append(*h,x.(int)) }; func (h *IH) Pop() interface{} { o:=*h; n:=len(o); x:=o[n-1]; *h=o[:n-1]; return x }; func main() { h:=&IH{1,2,3}; heap.Init(h); heap.Fix(h,0); heap.Remove(h,1) }",
    compress_lzw_writer => "package main; import \"compress/lzw\"; import \"bytes\"; func main() { _ = lzw.NewWriter(bytes.NewBuffer(nil), lzw.LSB, 8) }",
    compress_lzw_reader => "package main; import \"compress/lzw\"; import \"bytes\"; func main() { _ = lzw.NewReader(bytes.NewReader(nil), lzw.LSB, 8) }",
    compress_bzip2_reader => "package main; import \"compress/bzip2\"; import \"bytes\"; func main() { _ = bzip2.NewReader(bytes.NewReader(nil)) }",
    text_scanner_scan => "package main; import \"text/scanner\"; import \"strings\"; func main() { var s scanner.Scanner; s.Init(strings.NewReader(\"a\")); _, _, _ = s.Scan() }",
    regexp_syntax_parse => "package main; import \"regexp/syntax\"; func main() { _, _ = syntax.Parse(\"a+\", syntax.Perl) }",
    regexp_syntax_compile => "package main; import \"regexp/syntax\"; func main() { re, _ := syntax.Parse(\"a\", syntax.Perl); _ = syntax.Compile(re) }",
    io_copybuffer => "package main; import \"io\"; import \"strings\"; import \"bytes\"; func main() { _, _ = io.CopyBuffer(bytes.NewBuffer(nil), strings.NewReader(\"a\"), make([]byte, 8)) }",
    io_multireader => "package main; import \"io\"; import \"strings\"; func main() { _ = io.MultiReader(strings.NewReader(\"a\"), strings.NewReader(\"b\")) }",
    bufio_readslice => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a\\nb\")); _, _ = r.ReadSlice('\\n') }",
    bufio_peek => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"ab\")); _, _ = r.Peek(1) }",
}
