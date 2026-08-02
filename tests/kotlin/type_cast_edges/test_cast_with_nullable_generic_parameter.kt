// vybe-test: kotlin/type_cast_edges/test_cast_with_nullable_generic_parameter
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

class Box<T>(val value: T)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Box("x")
            val cast = value as? Box<String>
            __check((cast?.value ?: "none").toString(), "x")
        }
