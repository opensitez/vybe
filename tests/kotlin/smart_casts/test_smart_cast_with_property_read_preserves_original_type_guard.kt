// vybe-test: kotlin/smart_casts/test_smart_cast_with_property_read_preserves_original_type_guard
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

class Holder {
            val value: String = "v"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Holder()
            if (value is Holder) {
                __check((value.value).toString(), "v")
            }
            __check((value is Holder).toString(), "true")
        }
