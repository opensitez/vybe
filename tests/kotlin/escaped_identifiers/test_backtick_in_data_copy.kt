// vybe-test: kotlin/escaped_identifiers/test_backtick_in_data_copy
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

data class `Pair Data`(val `left value`: Int, val `right value`: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val p = `Pair Data`(1, 2)
    val q = p.copy(`right value` = 3)
    __check((q.`left value` + q.`right value`).toString(), "4")
}
