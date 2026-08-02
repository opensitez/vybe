// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            val map = values.associateBy({ it }, { it.toInt() - 96 })
            __check((map["a"]).toString(), "97")
            __check((map["c"]).toString(), "99")
        }
