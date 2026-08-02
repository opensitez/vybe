// vybe-test: kotlin/data_classes/test_data_class_nested_copy_propagates_outer
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Child(val value: Int)
        data class Parent(val child: Child, val tag: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p1 = Parent(Child(1), "x")
            val p2 = p1.copy(child = Child(9))
            __check((p1.child.value).toString(), "1")
            __check((p2.child.value).toString(), "9")
            __check((p2.tag).toString(), "x")
        }
