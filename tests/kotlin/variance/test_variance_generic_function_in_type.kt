// vybe-test: kotlin/variance/test_variance_generic_function_in_type
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Writer<in T> {
            fun write(value: T)
        }
        open class Thing
        class ThingWriter : Writer<Thing> {
            override fun write(value: Thing) { __check(("w").toString(), "w") }
        }
        fun consumeWriter(writer: Writer<Thing>) {
            writer.write(Thing())
        }
        class Fancy : Thing()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val writer: Writer<Fancy> = ThingWriter()
            consumeWriter(writer)
        }
