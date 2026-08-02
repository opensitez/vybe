// vybe-test: kotlin/data_class_copying/test_data_class_copy_multiple_instances
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Row(val id: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Row(1)
            val second = first.copy(2)
            val third = second.copy(3)
            __check((first.id).toString(), "1")
            __check((second.id).toString(), "2")
            __check((third.id).toString(), "3")
        }
