// vybe-test: kotlin/variance/test_variance_transform_out_to_in
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Producer<out T> { fun get(): T }
        interface Consumer<in T> { fun put(v: T) }
        class StringProducer : Producer<String> {
            override fun get(): String = "ok"
        }
        class Printer : Consumer<Any> {
            override fun put(v: Any) { __p((v.toString()).toString()) }
        }
        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val producer: Producer<Any> = StringProducer()
            val consumer: Consumer<String> = Printer()
            consumer.put(producer.get())
        
__check("ok")
}
