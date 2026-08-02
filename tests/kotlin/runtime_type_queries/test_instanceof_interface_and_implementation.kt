// vybe-test: kotlin/runtime_type_queries/test_instanceof_interface_and_implementation
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            interface Marker
            class A : Marker
            class B
            val a: Marker = A()
            val b: Any = B()
            __check((a is Marker).toString(), "true")
            __check((b is Marker).toString(), "false")
            __check((a !is B).toString(), "true")
            __check((b is B).toString(), "true")
        }
