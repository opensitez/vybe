// vybe-test: kotlin/equality_hashcode/test_to_string_for_data_class_is_stable
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Entry(val key: String, val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Entry("a", 3).toString()).toString(), "Entry(key=a, value=3)")
        }
