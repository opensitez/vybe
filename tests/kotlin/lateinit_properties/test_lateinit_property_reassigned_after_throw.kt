// vybe-test: kotlin/lateinit_properties/test_lateinit_property_reassigned_after_throw
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Holder {
            lateinit var value: String

            fun first() { value = "v1" }
            fun second() { value = "v2" }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            val a = try {
                h.value
                "ok"
            } catch (e: UninitializedPropertyAccessException) {
                "first-failed"
            }
            h.first()
            __check((a).toString(), "first-failed")
            h.second()
            __check((h.value).toString(), "v2")
        }
