// vybe-test: kotlin/secondary_constructors/test_secondary_chain_with_defaults
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Packet {
            val x: Int

            constructor() {
                this.x = 0
            }

            constructor(v: Int) : this() {
                this.x = v
            }

            constructor(v: Int, d: Int, e: Int) : this(v) {
                this.x = v + d + e
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Packet(2, 3, 4)
            __check((p.x).toString(), "9")
        }
