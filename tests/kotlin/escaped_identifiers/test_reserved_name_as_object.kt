// vybe-test: kotlin/escaped_identifiers/test_reserved_name_as_object
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val `val` = 3
        val `var` = 4
        __check((`val` + `var`).toString(), "7")
    }
