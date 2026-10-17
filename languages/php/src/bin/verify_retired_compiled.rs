fn main() {
    assert!(cfg!(debug_assertions), "service verification requires DEBUG");
    vybe_language_php::register();
    vybe_compiler::dynamic::verify_retired_compiled_reuse();
    println!("4 retired-compiled service checks passed");
}
