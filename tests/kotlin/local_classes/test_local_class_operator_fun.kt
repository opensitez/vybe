// vybe-test: kotlin/local_classes/test_local_class_operator_fun
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Box(val v: Int) {
                operator fun plus(other: Box) = Box(v + other.v)
            }
            __check(((Box(2) + Box(3)).v).toString(), "5")
        }
