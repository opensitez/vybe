use crate::helpers::run_prints;

#[test]
fn test_generic_extension_with_type_inference() {
    let out = run_prints(r#"
        fun <T : Any> T?.orFallback(default: T): T = this ?: default

        fun main() {
            val text: String? = null
            val count: Int? = 4
            println(text.orFallback("x"))
            println(count.orFallback(9))
        }
    "#);
    assert_eq!(out, &["x", "4"]);
}

#[test]
fn test_extension_generic_infix_pairing() {
    let out = run_prints(r#"
        fun <T> T.thenValue(v: T): List<T> = listOf(this, v)

        fun main() {
            println(1.thenValue(2))
            println("a".thenValue("b"))
        }
    "#);
    assert_eq!(out, &["[1, 2]", "[a, b]"]);
}
