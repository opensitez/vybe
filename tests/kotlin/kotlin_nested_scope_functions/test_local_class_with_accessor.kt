// vybe-test: kotlin/kotlin_nested_scope_functions/test_local_class_with_accessor
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local(val base: Int) {
                fun value() = base * 2
            }
            val local = Local(4)
            __check((local.value()).toString(), "8")
        }
