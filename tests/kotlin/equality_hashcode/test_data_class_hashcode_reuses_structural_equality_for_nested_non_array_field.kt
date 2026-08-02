// vybe-test: kotlin/equality_hashcode/test_data_class_hashcode_reuses_structural_equality_for_nested_non_array_field
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Child(val token: String)
        data class Parent(val child: Child)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Parent(Child("x"))
            val right = Parent(Child("x"))
            __check((left == right).toString(), "true")
            __check((left.hashCode() == right.hashCode()).toString(), "true")
        }
