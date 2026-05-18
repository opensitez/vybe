fn main() {
    let src = "package main; import \"fmt\"; func main() { fmt.Println(0xFF); }";
    let ast = vybe_compiler::languages::go::parse(src).unwrap();
    println!("{:#?}", ast);
}
