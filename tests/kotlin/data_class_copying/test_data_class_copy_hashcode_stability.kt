// vybe-test: kotlin/data_class_copying/test_data_class_copy_hashcode_stability
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Entry(val key: String, val score: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Entry("x", 1)
            val b = a.copy()
            __check((a.hashCode() == b.hashCode()).toString(), "true")
            __check((a == b).toString(), "true")
        }
