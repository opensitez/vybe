// vybe-test: kotlin/classes/test_class_instance_init_order_with_multiple_inits
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class Base {
            init {
                __check(("base").toString(), "base")
            }
            open val baseName: String = "base"
        }

        class Child : Base() {
            init {
                __check(("child").toString(), "child")
            }
            override val baseName: String = "child"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Child()
            __check((c.baseName).toString(), "child")
        }
