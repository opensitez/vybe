// vybe-test: kotlin/type_cast_edges/test_cast_to_superclass_interface
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

interface X
        class Y : X
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Y()
            __check((value is X).toString(), "true")
            val cast = value as X
            __check((cast is X).toString(), "true")
        }
