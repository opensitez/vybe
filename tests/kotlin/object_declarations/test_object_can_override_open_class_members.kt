// vybe-test: kotlin/object_declarations/test_object_can_override_open_class_members
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

open class Logger {
            open fun level(): String = "base"
        }

        object Runtime : Logger() {
            override fun level(): String = "runtime"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Runtime.level()).toString(), "runtime")
        }
