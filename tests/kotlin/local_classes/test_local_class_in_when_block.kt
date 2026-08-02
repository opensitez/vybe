// vybe-test: kotlin/local_classes/test_local_class_in_when_block
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 1
            val out = when (n) {
                1 -> {
                    class L(val v: String)
                    L("a").v
                }
                else -> "z"
            }
            __check((out).toString(), "a")
        }
