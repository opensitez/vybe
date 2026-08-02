// vybe-test: kotlin/companion_objects/test_companion_object_methods_can_return_receiver_instance
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder {
            val marker: String
            private constructor(marker: String) {
                this.marker = marker
            }

            companion object {
                fun create(): Holder = Holder("ok")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder.create().marker).toString(), "ok")
        }
