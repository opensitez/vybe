// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_infer_from_assignment
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

class Box<T>(val value: T)

        fun <T> firstOrDefault(value: T?): T {
            return value ?: throw IllegalStateException("missing")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: String? = "k"
            val b = Box(firstOrDefault(v))
            __check((b.value).toString(), "k")
        }
