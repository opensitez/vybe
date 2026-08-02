// vybe-test: kotlin/kotlin_associate_apis/test_associate_from_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = sequenceOf("a", "bb", "ccc").associateBy { it.length }
            __check((map[1]).toString(), "a")
            __check((map[3]).toString(), "ccc")
        }
