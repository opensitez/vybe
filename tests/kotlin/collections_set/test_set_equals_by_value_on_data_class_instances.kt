// vybe-test: kotlin/collections_set/test_set_equals_by_value_on_data_class_instances
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Label(val id: Int, val name: String)
            val values = setOf(Label(1, "x"))
            __check((values.contains(Label(1, "x"))).toString(), "true")
            __check((values.contains(Label(1, "y"))).toString(), "false")
        }
