// vybe-test: kotlin/data_classes/test_data_class_named_component_access
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class PairValue(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = PairValue(4, 9)
            __check((p.component1()).toString(), "4")
            __check((p.component2()).toString(), "9")
        }
