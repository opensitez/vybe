// vybe-test: kotlin/kotlin_uuid/test_uuid_variant_and_version_bits
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val id = java.util.UUID.fromString("ffffffff-ffff-4fff-8fff-ffffffffffff")
            __check((id.version()).toString(), "4")
            __check((id.variant()).toString(), "2")
            __check((id.variant() == 2).toString(), "true")
        }
