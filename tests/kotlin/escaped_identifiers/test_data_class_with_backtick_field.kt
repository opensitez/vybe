// vybe-test: kotlin/escaped_identifiers/test_data_class_with_backtick_field
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

data class `Item Box`(val `item id`: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val i = `Item Box`(4)
__check((i.`item id`).toString(), "4") }
