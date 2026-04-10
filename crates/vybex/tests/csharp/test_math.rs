use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn math_floor() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Floor(3.7));"), "3");
}

#[test]
fn math_abs() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Abs(-5));"), "5");
}

#[test]
fn math_sqrt() {
    assert_eq!(run_csharp_one("Console.WriteLine(Math.Sqrt(16));"), "4");
}

#[test]
fn math_multiple() {
    let out = run_csharp(r#"
        Console.WriteLine(Math.Floor(9.7));
        Console.WriteLine(Math.Abs(-42));
        Console.WriteLine(Math.Sqrt(144));
    "#);
    assert_eq!(out, vec!["9", "42", "12"]);
}
