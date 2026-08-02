// vybe-test: kotlin/inner_classes/test_inner_class_method_calls_outer
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Counter(val prefix: String) {
            inner class Marker {
                fun label(v: Int): String = "$prefix-$v"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter("k").Marker().label(9)).toString(), "k-9")
        }
