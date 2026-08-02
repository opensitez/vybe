// vybe-test: kotlin/collections/test_empty_array_behavior
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = arrayOf<Int>()
            __check((empty.size).toString(), "0")
            if (empty.size == 0) {
                __check(("empty").toString(), "empty")
            }
        }
