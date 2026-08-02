// vybe-test: kotlin/local_classes/test_local_class_in_if_block
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = true
            if (x) {
                class Local(val v: Int)
                __check((Local(7).v).toString(), "7")
            }
        }
