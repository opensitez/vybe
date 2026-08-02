// vybe-test: kotlin/data_classes/test_data_class_in_set_with_mutated_var_field_can_change_lookup
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Packet(var id: Int, val payload: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = Packet(1, "x")
            val set = mutableSetOf(one)
            __check((set.contains(Packet(1, "x"))).toString(), "true")
            one.id = 2
            __check((set.contains(Packet(2, "x"))).toString(), "false")
            __check((set.contains(Packet(1, "x"))).toString(), "false")
        }
