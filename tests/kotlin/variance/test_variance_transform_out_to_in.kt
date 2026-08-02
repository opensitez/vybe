// vybe-test: kotlin/variance/test_variance_transform_out_to_in
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Producer<out T> { fun get(): T }
        interface Consumer<in T> { fun put(v: T) }
        class StringProducer : Producer<String> {
            override fun get(): String = "ok"
        }
        class Printer : Consumer<Any> {
            override fun put(v: Any) { __check((v.toString()).toString(), "ok") }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val producer: Producer<Any> = StringProducer()
            val consumer: Consumer<String> = Printer()
            consumer.put(producer.get())
        }
