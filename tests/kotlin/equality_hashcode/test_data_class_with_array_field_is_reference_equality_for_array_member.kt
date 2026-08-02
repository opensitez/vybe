// vybe-test: kotlin/equality_hashcode/test_data_class_with_array_field_is_reference_equality_for_array_member
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Holder(val values: Array<Int>)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Holder(arrayOf(1, 2, 3))
            val second = Holder(arrayOf(1, 2, 3))
            val third = Holder(first.values)
            __check((first == second).toString(), "false")
            __check((first == third).toString(), "false")
            __check((first.values === third.values).toString(), "true")
        }
