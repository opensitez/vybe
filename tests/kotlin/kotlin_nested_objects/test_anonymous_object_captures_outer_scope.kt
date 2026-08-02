// vybe-test: kotlin/kotlin_nested_objects/test_anonymous_object_captures_outer_scope
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_objects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val label = "a"
            val obj = object {
                fun render(): String = label + "b"
            }
            __check((obj.render()).toString(), "ab")
        }
