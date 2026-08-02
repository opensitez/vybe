// vybe-test: kotlin/object_declarations/test_object_can_be_extended_from_open_class
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

open class Base {
            open fun tag(): String = "base"
        }

        object Child : Base() {
            override fun tag(): String = "child"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child.tag()).toString(), "child")
        }
