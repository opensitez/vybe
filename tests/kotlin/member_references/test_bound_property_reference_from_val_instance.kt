// vybe-test: kotlin/member_references/test_bound_property_reference_from_val_instance
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Counter {
            val label: String = "c"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counter = Counter()
            val labelRef = counter::label
            __check((labelRef()).toString(), "c")
        }
