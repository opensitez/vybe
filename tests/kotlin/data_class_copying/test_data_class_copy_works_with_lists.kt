// vybe-test: kotlin/data_class_copying/test_data_class_copy_works_with_lists
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Bucket(val items: List<Int>)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Bucket(listOf(1, 2))
            val changed = base.copy(items = base.items + listOf(3, 4))
            __check((base.items.joinToString(",")).toString(), "1,2")
            __check((changed.items.joinToString(",")).toString(), "1,2,3,4")
        }
