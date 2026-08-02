// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_boolean_key
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf(1, 2, 3, 4).associateBy { it % 2 == 0 }
            __check((map.keys.joinToString(",")).toString(), "false,true")
            __check((map[true]?.size).toString(), "2")
        }
