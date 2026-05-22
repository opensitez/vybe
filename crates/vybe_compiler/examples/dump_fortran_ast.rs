use std::env;
use std::fs;

fn main() {
    let path = env::args().nth(1).expect("expected source path");
    let source = fs::read_to_string(&path).expect("read source");
    let module = vybe_compiler::languages::fortran::parse(&source).expect("parse fortran");
    println!("{:#?}", module);
}
