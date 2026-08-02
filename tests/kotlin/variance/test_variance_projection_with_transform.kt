// vybe-test: kotlin/variance/test_variance_projection_with_transform
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun stringify(values: List<out Any?>): String = values.joinToString(",")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((stringify(listOf(1, "a", true))).toString(), "1,a,true")
        }
