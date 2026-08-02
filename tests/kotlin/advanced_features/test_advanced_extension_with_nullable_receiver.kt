// vybe-test: kotlin/advanced_features/test_advanced_extension_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Holder(val value: Int)

        fun Holder.incremented(): Holder {
            return Holder(this.value + 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder(10)
            __check((h.incremented().value).toString(), "11")
        }
