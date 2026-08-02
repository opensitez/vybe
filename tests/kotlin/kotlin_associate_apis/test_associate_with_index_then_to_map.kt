// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_index_then_to_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = (0..3).associateWith { it * 2 }
            __check((map[2]).toString(), "4")
            __check((map.values.sum()).toString(), "12")
        }
