// vybe-test: kotlin/companion_objects/test_generic_companion_factory_preserves_inferred_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder<T>(val value: T) {
            companion object {
                fun <T> make(value: T): Holder<T> = Holder(value)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = Holder.make("kotlin").value
            val number = Holder.make(12).value
            __check((text).toString(), "kotlin")
            __check((number).toString(), "12")
        }
