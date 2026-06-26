//! ast and token compile-only (go/parser, go/token) patterns.

use crate::helpers::*;

go_compile_cases! {
    token_file_set => "package main; import \"go/token\"; func main() { fs := token.NewFileSet(); _ = fs }",
    token_file_pos => "package main; import \"go/token\"; func main() { fs := token.NewFileSet(); f := fs.AddFile(\"a.go\", fs.Base(), 100); _ = f.Pos(1) }",
    parser_parse_file => "package main; import \"go/parser\"; import \"go/token\"; func main() { fs := token.NewFileSet(); _, _ = parser.ParseFile(fs, \"\", \"package main\\nfunc main() {}\", 0) }",
    parser_parse_expr => "package main; import \"go/parser\"; import \"go/token\"; func main() { _, _ = parser.ParseExpr(\"1+2\") }",
    parser_parse_dir => "package main; import \"go/parser\"; import \"go/token\"; func main() { fs := token.NewFileSet(); _, _ = parser.ParseDir(fs, \".\", nil, 0) }",
    ast_walk => "package main; import \"go/ast\"; import \"go/parser\"; import \"go/token\"; func main() { fs := token.NewFileSet(); f, _ := parser.ParseFile(fs, \"\", \"package main\", 0); ast.Walk(nil, f) }",
    ast_filter => "package main; import \"go/ast\"; import \"go/parser\"; func main() { f, _ := parser.ParseFile(nil, \"\", \"package main\", 0); ast.FilterFile(f, func(s string) bool { return true }) }",
    ast_comment_map => "package main; import \"go/ast\"; import \"go/parser\"; import \"go/token\"; func main() { fs := token.NewFileSet(); f, _ := parser.ParseFile(fs, \"\", \"package main // c\", parser.ParseComments); _ = ast.NewCommentMap(fs, f, f.Comments) }",
    scanner_init => "package main; import \"go/scanner\"; import \"go/token\"; func main() { var s scanner.Scanner; fs := token.NewFileSet(); f := fs.AddFile(\"x.go\", fs.Base(), 10); s.Init(f, []byte(\"package main\"), nil, 0) }",
    scanner_scan => "package main; import \"go/scanner\"; import \"go/token\"; func main() { var s scanner.Scanner; fs := token.NewFileSet(); f := fs.AddFile(\"x.go\", fs.Base(), 20); s.Init(f, []byte(\"package main\"), nil, scanner.ScanComments); _, _, _ = s.Scan() }",
}
