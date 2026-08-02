// vybe-test: kotlin/kotlin_nested_objects/test_nested_named_object_is_accessible
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_objects.rs

object Container {
            object Tag {
                fun value(): String = "nested"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Container.Tag.value()).toString(), "nested")
        }
