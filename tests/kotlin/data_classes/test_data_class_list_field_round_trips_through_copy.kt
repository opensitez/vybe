// vybe-test: kotlin/data_classes/test_data_class_list_field_round_trips_through_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Bucket(val values: MutableList<Int>)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Bucket(mutableListOf(1, 2))
            val copy = base.copy()
            copy.values.add(3)
            __check((base.values.size).toString(), "3")
            __check((copy.values.size).toString(), "3")
            __check((base.values[2]).toString(), "3")
        }
