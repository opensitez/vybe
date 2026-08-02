// vybe-test: kotlin/nested_classes/test_nested_class_in_function_scope
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

fun factory(): String {
            class Packet(val value: Int)
            return Packet(3).value.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((factory()).toString(), "3")
        }
