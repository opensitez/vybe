// vybe-test: kotlin/classes/test_nested_class
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Outer {
            class Nested {
                fun getMsg(): String = "nested msg"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = Outer.Nested()
            __check((n.getMsg()).toString(), "nested msg")
        }
