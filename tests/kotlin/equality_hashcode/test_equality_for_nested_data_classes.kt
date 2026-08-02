// vybe-test: kotlin/equality_hashcode/test_equality_for_nested_data_classes
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Child(val value: Int)
        data class Parent(val child: Child)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Parent(Child(1))
            val b = Parent(Child(1))
            __check((a == b).toString(), "true")
            __check((a.child == b.child).toString(), "true")
            __check((a.child === b.child).toString(), "false")
        }
