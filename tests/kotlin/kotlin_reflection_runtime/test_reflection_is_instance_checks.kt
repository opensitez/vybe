// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_is_instance_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = ProbeImpl(1)
            __check((Probe::class.isInstance(value)).toString(), "true")
            __check((MarkerContract::class.isInstance(value)).toString(), "true")
            __check((Probe::class.isInstance("x")).toString(), "false")
            __check((ProbeImpl::class.isInstance(value)).toString(), "true")
        }
