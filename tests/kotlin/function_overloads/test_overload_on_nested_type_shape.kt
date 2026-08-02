// vybe-test: kotlin/function_overloads/test_overload_on_nested_type_shape
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun parse(v: Int): String = "int"
        fun parse(v: Any): String = "any"
        open class A
        class B : A()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((parse(1)).toString(), "int")
            __check((parse(B())).toString(), "any")
        }
