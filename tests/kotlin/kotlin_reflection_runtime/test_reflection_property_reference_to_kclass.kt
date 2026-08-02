// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_property_reference_to_kclass
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ref = Probe::class
            __check((ref.isInstance(Probe("id"))).toString(), "true")
            __check((ref.isInstance(123)).toString(), "false")
            __check((ref.toString().contains("KClass")).toString(), "true")
        }
