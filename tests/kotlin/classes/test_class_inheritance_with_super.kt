// vybe-test: kotlin/classes/test_class_inheritance_with_super
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class Parent {
            open fun name(): String = "parent"
        }

        class Child : Parent() {
            override fun name(): String {
                return super.name() + "-child"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            __check((child.name()).toString(), "parent-child")
        }
