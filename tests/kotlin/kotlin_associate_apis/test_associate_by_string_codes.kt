// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_string_codes
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("a", "b", "c").associateBy({ it.first() }, { it.code })
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map["a"]).toString(), "97")
        }
