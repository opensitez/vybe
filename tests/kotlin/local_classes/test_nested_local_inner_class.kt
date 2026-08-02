// vybe-test: kotlin/local_classes/test_nested_local_inner_class
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Outer {
                inner class Inner(val base: String)
            }
            val o = Outer()
            __check((o.Inner("x").base).toString(), "x")
        }
