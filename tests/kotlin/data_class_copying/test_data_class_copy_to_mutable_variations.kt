// vybe-test: kotlin/data_class_copying/test_data_class_copy_to_mutable_variations
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Counter(val values: MutableList<Int>)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Counter(mutableListOf(1, 2))
            val b = a.copy()
            b.values.add(3)
            __check((a.values.joinToString(",")).toString(), "1,2,3")
            __check((b.values.joinToString(",")).toString(), "1,2,3")
        }
