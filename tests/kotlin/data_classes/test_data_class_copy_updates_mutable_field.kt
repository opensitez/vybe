// vybe-test: kotlin/data_classes/test_data_class_copy_updates_mutable_field
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Counter(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter(2)
            val d = c.copy(value = 12)
            __check((c.value).toString(), "2")
            __check((d.value).toString(), "12")
            d.value += 1
            __check((d.value).toString(), "13")
        }
