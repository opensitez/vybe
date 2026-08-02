// vybe-test: kotlin/extension_properties/test_extension_property_nullable_string_or_empty
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String?.valueOrDash: String get() = this ?: "-"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: String? = null
            val b: String? = "ok"
            __check((a.valueOrDash).toString(), "-")
            __check((b.valueOrDash).toString(), "ok")
        }
