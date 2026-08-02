// vybe-test: kotlin/local_classes/test_local_class_operator_plus
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class L(val v: Int) {
                operator fun inc() = L(v + 1)
            }
            __check(((L(1).inc().v)).toString(), "2")
        }
