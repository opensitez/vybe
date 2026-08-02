// vybe-test: kotlin/local_classes/test_nested_local_class
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Outer {
                class Inner(val v: Int)
            }
            __check((Outer.Inner(1).v).toString(), "1")
        }
