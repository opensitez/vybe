use crate::helpers::run_prints;

#[test]
fn test_anonymous_object_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val runner = object {
                fun run() {
                    println("anonymous object running")
                }
            }
            runner.run()
        }
    "#,
    );
    assert_eq!(out, &["anonymous object running"]);
}

#[test]
fn test_object_expression_with_interface() {
    let out = run_prints(
        r#"
        interface Callback {
            fun onComplete()
        }

        fun main() {
            val cb = object : Callback {
                override fun onComplete() {
                    println("callback finished")
                }
            }
            cb.onComplete()
        }
    "#,
    );
    assert_eq!(out, &["callback finished"]);
}

#[test]
fn test_object_expression_with_state() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = object {
                var value = 0
                fun inc() {
                    value += 1
                }
                fun reset() {
                    value = 0
                }
            }

            counter.inc()
            counter.inc()
            counter.reset()
            counter.inc()
            println(counter.value)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_anonymous_object_with_init_and_state() {
    let out = run_prints(
        r#"
        fun main() {
            val stateful = object {
                var value = 0
                fun inc() {
                    value += 1
                }
                init {
                    value = 5
                }
            }

            stateful.inc()
            println(stateful.value)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_object_expression_as_typed_interface() {
    let out = run_prints(
        r#"
        interface Worker {
            fun work(): String
        }

        fun main() {
            val w: Worker = object : Worker {
                override fun work(): String {
                    return "done"
                }
            }
            println(w.work())
        }
    "#,
    );
    assert_eq!(out, &["done"]);
}

#[test]
fn test_object_expression_stored_and_reused() {
    let out = run_prints(
        r#"
        fun main() {
            val provider = object {
                fun value(): Int {
                    return 3
                }
            }

            val first = provider.value()
            val second = provider.value()
            println(first + second)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_object_expression_with_arguments_like_member() {
    let out = run_prints(
        r#"
        fun main() {
            val adder = object {
                fun apply(base: Int, extra: Int): Int {
                    return base + extra
                }
            }
            println(adder.apply(2, 5))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_object_expression_property_access() {
    let out = run_prints(
        r#"
        fun main() {
            val record = object {
                val id = 10
                val label = "log"
            }
            println(record.id)
            println(record.label)
        }
    "#,
    );
    assert_eq!(out, &["10", "log"]);
}

#[test]
fn test_object_expression_mutating_counter() {
    let out = run_prints(
        r#"
        fun main() {
            val counter = object {
                var value = 1
                fun inc() {
                    value *= 2
                }
            }

            counter.inc()
            counter.inc()
            println(counter.value)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_object_expression_implements_multiple_methods() {
    let out = run_prints(
        r#"
        interface A {
            fun a(): Int
        }

        interface B {
            fun b(): Int
        }

        fun main() {
            val combined = object : A, B {
                override fun a(): Int { return 1 }
                override fun b(): Int { return 2 }
            }
            println(combined.a() + combined.b())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_object_expression_nested_creation() {
    let out = run_prints(
        r#"
        fun main() {
            val builder = object {
                fun create(): String {
                    return "x"
                }
            }

            val name = builder.create()
            val wrapper = object {
                fun wrap(value: String): String {
                    return value + value
                }
            }

            println(wrapper.wrap(name))
        }
    "#,
    );
    assert_eq!(out, &["xx"]);
}

#[test]
fn test_object_expression_with_if() {
    let out = run_prints(
        r#"
        fun main() {
            val obj = object {
                fun flag(x: Int): String {
                    if (x > 0) {
                        return "yes"
                    }
                    return "no"
                }
            }
            println(obj.flag(1))
            println(obj.flag(-1))
        }
    "#,
    );
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_object_expression_as_returned_value() {
    let out = run_prints(
        r#"
        interface Producer {
            fun value(): Int
        }

        fun makeProducer(start: Int): Producer {
            return object : Producer {
                override fun value(): Int {
                    return start
                }
            }
        }

        fun main() {
            val p = makeProducer(9)
            println(p.value())
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_object_expression_return_string() {
    let out = run_prints(
        r#"
fun makeLabel(): String { val obj = object { fun text() = "ok" }; return obj.text() }; fun main() { println(makeLabel()) }
"#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_object_expression_two_methods() {
    let out = run_prints(
        r#"
interface MathOp { fun a(): Int; fun b(): Int }; fun makeOps() = object : MathOp { override fun a() = 2; override fun b() = 3 }; fun main() { val op = makeOps(); println(op.a() + op.b()) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_expression_property_and_method() {
    let out = run_prints(
        r#"
fun main() { val obj = object { var value = 1; fun inc() { value += 1 } }; obj.inc(); obj.inc(); println(obj.value) }
"#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_object_expression_as_argument() {
    let out = run_prints(
        r#"
interface Sink { fun consume(v: Int) }; fun call(sink: Sink, value: Int) = sink.consume(value); fun main() { call(object : Sink { override fun consume(v: Int) { println(v) } }, 7) }
"#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_object_expression_inside_loop() {
    let out = run_prints(
        r#"
fun main() { val obj = object { var value = 0; fun inc() { value += 1 } }; for (i in 1..3) { obj.inc() }; println(obj.value) }
"#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_object_expression_with_state_reset() {
    let out = run_prints(
        r#"
fun main() { val obj = object { var value = 0; fun add(v: Int) { value += v }; fun reset() { value = 0 } }; obj.add(5); obj.reset(); println(obj.value) }
"#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_object_expression_from_function() {
    let out = run_prints(
        r#"
fun create(): Int { val worker = object { fun run(v: Int) = v + 1 }; return worker.run(4) }; fun main() { println(create()) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_expression_in_while() {
    let out = run_prints(
        r#"
fun main() { val o = object { var n = 0; fun next() { n += 1 } }; var i = 0; while (i < 2) { o.next(); i += 1 }; println(o.n) }
"#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_object_expression_as_predicate() {
    let out = run_prints(
        r#"
interface Check { fun ok(v: Int): Boolean }; fun runCheck(c: Check): Boolean = c.ok(5); fun main() { println(runCheck(object : Check { override fun ok(v: Int) = v > 3 })) }
"#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_object_expression_with_return() {
    let out = run_prints(
        r#"
fun makeOutput() = object { fun out() = "pong" }; fun main() { val o = makeOutput(); println(o.out()) }
"#,
    );
    assert_eq!(out, &["pong"]);
}

#[test]
fn test_object_expression_double_field() {
    let out = run_prints(
        r#"
fun main() { val pair = object { var left = 1; var right = 2 }; pair.left += 3; println(pair.left + pair.right) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_object_expression_chain() {
    let out = run_prints(
        r#"
fun main() { val a = object { fun first() = 2 }; val b = object { fun second() = 3 }; println(a.first() + b.second()) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_expression_with_args() {
    let out = run_prints(
        r#"
fun main() { val sum = object { fun add(x: Int, y: Int) = x + y }; println(sum.add(4, 5)) }
"#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_object_expression_as_function() {
    let out = run_prints(
        r#"
fun make(): Int { val f = object { fun call(v: Int) = v * 3 }; return f.call(2) }; fun main() { println(make()) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_object_expression_boolean_case() {
    let out = run_prints(
        r#"
fun main() { val o = object { fun check(v: Int) = v % 2 == 0 }; println(o.check(2)); println(o.check(7)) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_object_expression_local_type() {
    let out = run_prints(
        r#"
fun main() { val result = object { val value = 1 }; println(result.value + 4) }
"#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_object_expression_with_mutable_interface_property() {
    let out = run_prints(
        r#"
interface Tally {
    fun next(): Int
    fun total(): Int
}

fun main() {
    val tally = object : Tally {
        var count = 0
        override fun next(): Int {
            count += 1
            return count
        }

        override fun total(): Int = count
    }

    println(tally.next())
    println(tally.next())
    println(tally.total())
}
"#,
    );
    assert_eq!(out, &["1", "2", "2"]);
}

#[test]
fn test_object_expression_can_extend_open_class() {
    let out = run_prints(
        r#"
        open class Base {
            open fun label(): String = "base"
        }

        fun main() {
            val value = object : Base() {
                override fun label(): String {
                    return super.label() + "-child"
                }
            }
            println(value.label())
        }
    "#,
    );
    assert_eq!(out, &["base-child"]);
}

#[test]
fn test_object_expression_capture_outer_mutable_var() {
    let out = run_prints(
        r#"
        fun main() {
            var prefix = "left"
            val obj = object {
                fun build(value: String): String = prefix + ":" + value
            }
            println(obj.build("a"))
            prefix = "right"
            println(obj.build("b"))
        }
    "#,
    );
    assert_eq!(out, &["left:a", "right:b"]);
}

#[test]
fn test_object_expression_with_custom_getter_and_setter() {
    let out = run_prints(
        r#"
        fun main() {
            val obj = object {
                var counter = 0
                val doubled: Int
                    get() = counter * 2
            }
            obj.counter = 5
            println(obj.doubled)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_object_expression_as_nested_return_value() {
    let out = run_prints(
        r#"
        interface Calculator {
            fun add(value: Int): Int
        }

        fun wrap(base: Int): Calculator {
            return object : Calculator {
                override fun add(value: Int): Int {
                    return base + value
                }
            }
        }

        fun main() {
            val calc = wrap(4)
            println(calc.add(3))
            println(calc.add(1))
        }
    "#,
    );
    assert_eq!(out, &["7", "5"]);
}

#[test]
fn test_object_expression_and_outer_this_in_member_scope() {
    let out = run_prints(
        r#"
        class Envelope(val marker: String) {
            fun make(): String {
                val obj = object {
                    fun value(): String = this@Envelope.marker
                }
                return obj.value()
            }
        }

        fun main() {
            println(Envelope("ok").make())
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_object_expression_implements_interface_via_type_alias() {
    let out = run_prints(
        r#"
        interface Reader {
            fun read(): String
            fun fallback(): String = "none"
        }

        fun main() {
            val reader: Reader = object : Reader {
                override fun read(): String = "ok"
            }
            println(reader.read())
            println(reader.fallback())
        }
    "#,
    );
    assert_eq!(out, &["ok", "none"]);
}
