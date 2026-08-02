// vybe-test: kotlin/kotlin_associate_apis/test_associate_projection_pairs
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf(1 to "one", 2 to "two")
            val keys = map.keys.associateWith { it * 10 }
            __check((keys[1]).toString(), "10")
            __check((keys[2]).toString(), "20")
        }
