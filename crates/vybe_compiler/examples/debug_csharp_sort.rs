use vybe_compiler::compiler::Compiler;
use vybe_compiler::languages::csharp;
use vybe_compiler::profile::parse_profile;

fn main() {
    let src = r#"
using System.Collections.Generic;
class Temperature : IComparable<Temperature> {
    public double Degrees;
    public Temperature(double d) { Degrees = d; }
    public int CompareTo(Temperature other) {
        return Degrees.CompareTo(other.Degrees);
    }
    public override string ToString() { return Degrees + "°"; }
}
var temps = new List<Temperature> {
    new Temperature(100),
    new Temperature(37),
    new Temperature(0)
};
temps.Sort();
foreach (var t in temps) Console.WriteLine(t);
"#;

    let module = csharp::parse(src).expect("parse");
    let profile = parse_profile(csharp::profile_source()).expect("profile");
    let chunks = Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");

    for (index, chunk) in chunks.iter().enumerate() {
        println!("\n-- chunk {index}: {} --", chunk.name);
        println!("{}", vybe_bytecode::debug::disassemble(chunk));
        if !chunk.constants.is_empty() {
            println!("constants:");
            for (constant_index, constant) in chunk.constants.iter().enumerate() {
                println!("  [{constant_index}] {constant}");
            }
        }
    }
}
