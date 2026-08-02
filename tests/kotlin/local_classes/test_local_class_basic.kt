// vybe-test: kotlin/local_classes/test_local_class_basic
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local(val v: Int)
            __check((Local(1).v).toString(), "1")
        }
