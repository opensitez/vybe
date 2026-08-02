// vybe-test: kotlin/data_classes/test_data_class_var_mutation_changes_equality_result
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Box(var text: String, val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Box("a", 1)
            val b = Box("a", 1)
            val sameBefore = (a == b)
            a.text = "b"
            __check((sameBefore).toString(), "true")
            __check((a == b).toString(), "false")
            __check((a.text).toString(), "b")
        }
