// vybe-test: kotlin/equality_hashcode/test_data_class_to_string_includes_all_fields
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Entry(val id: Int, val label: String, val active: Boolean)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Entry(2, "x", true).toString()).toString(), "Entry(id=2, label=x, active=true)")
        }
