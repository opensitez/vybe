// vybe-test: kotlin/kotlin_annotation_usage/test_parameterized_local_annotation_is_parsed
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotation_usage.rs

@Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class Marker(val id: Int)

        fun compute(@Marker(42) value: Int): Int = value + 1

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compute(4)).toString(), "5")
        }
