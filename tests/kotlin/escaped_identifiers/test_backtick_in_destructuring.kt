// vybe-test: kotlin/escaped_identifiers/test_backtick_in_destructuring
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

data class `Node Pair`(val `a key`: Int, val `b key`: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val (`a key`, `b key`) = Pair(1, 2)
    __check((`a key` + `b key`).toString(), "3")
}
