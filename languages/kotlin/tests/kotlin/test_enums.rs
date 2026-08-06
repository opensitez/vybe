use crate::helpers::run_prints;

#[test]
fn test_enum_class_simple() {
    let out = run_prints(
        r#"
        enum class Direction {
            NORTH, SOUTH, EAST, WEST
        }

        fun main() {
            val dir = Direction.NORTH
            println(dir)
        }
    "#,
    );
    // Real Kotlin agrees: printing an enum entry gives its NAME
    // (`Enum.toString` is the name), never the ordinal.
    assert_eq!(out, &["NORTH"]);
}

#[test]
fn test_enum_class_matching() {
    let out = run_prints(
        r#"
        enum class Status {
            PENDING, APPROVED, REJECTED
        }

        fun main() {
            val s = Status.APPROVED
            // Was `s == 1` — kotlinc rejects comparing an enum to Int.
            if (s == Status.APPROVED) {
                println("Approved Status")
            } else {
                println("Other Status")
            }
        }
    "#,
    );
    assert_eq!(out, &["Approved Status"]);
}

#[test]
fn test_enum_all_entries() {
    let out = run_prints(
        r#"
        enum class Level { LOW, MEDIUM, HIGH }

        fun main() {
            println(Level.LOW)
            println(Level.MEDIUM)
            println(Level.HIGH)
        }
    "#,
    );
    // Real Kotlin agrees: names, not ordinals.
    assert_eq!(out, &["LOW", "MEDIUM", "HIGH"]);
}

#[test]
fn test_enum_equality_check() {
    let out = run_prints(
        r#"
        enum class State { OFF, ON }

        fun main() {
            val s1 = State.OFF
            val s2 = State.ON
            if (s1 != s2) {
                println("different states")
            }
        }
    "#,
    );
    assert_eq!(out, &["different states"]);
}

#[test]
fn test_enum_when_matching() {
    let out = run_prints(
        r#"
        enum class HttpStatus { OK, ERROR, UNKNOWN }

        fun describe(status: HttpStatus): String {
            return when (status) {
                HttpStatus.OK -> "ok"
                HttpStatus.ERROR -> "error"
                HttpStatus.UNKNOWN -> "unknown"
            }
        }

        fun main() {
            println(describe(HttpStatus.ERROR))
        }
    "#,
    );
    assert_eq!(out, &["error"]);
}

#[test]
fn test_enum_entry_with_payload() {
    let out = run_prints(
        r#"
        enum class Planet(val order: Int) {
            MERCURY(1),
            VENUS(2),
            EARTH(3)
        }

        fun main() {
            println(Planet.VENUS.order)
            println(Planet.EARTH.order + Planet.MERCURY.order)
        }
    "#,
    );
    assert_eq!(out, &["2", "4"]);
}

#[test]
fn test_enum_iteration() {
    let out = run_prints(
        r#"
        enum class Light { RED, GREEN, BLUE }

        fun main() {
            var count = 0
            for (entry in arrayOf(Light.RED, Light.GREEN, Light.BLUE)) {
                if (entry == Light.RED) {
                    count += 1
                } else if (entry == Light.GREEN) {
                    count += 1
                } else {
                    count += 1
                }
            }
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_enum_properties_and_accessors() {
    let out = run_prints(
        r#"
        enum class Level(val code: Int) {
            LOW(1),
            MEDIUM(2),
            HIGH(3)
        }

        fun main() {
            println(Level.MEDIUM.code)
            println(Level.HIGH.code)
        }
    "#,
    );
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_enum_value_equality_and_comparison() {
    let out = run_prints(
        r#"
        enum class Toggle { OFF, ON }

        fun main() {
            val a = Toggle.ON
            val b = Toggle.ON
            if (a == b) {
                println("same")
            }
        }
    "#,
    );
    assert_eq!(out, &["same"]);
}

#[test]
fn test_enum_in_when_with_range_conditions() {
    let out = run_prints(
        r#"
        enum class Grade { A, B, C, D }

        fun label(g: Grade): String {
            return when (g) {
                Grade.A -> "excellent"
                Grade.B -> "good"
                Grade.C -> "ok"
                Grade.D -> "need work"
            }
        }

        fun main() {
            println(label(Grade.A))
            println(label(Grade.D))
        }
    "#,
    );
    assert_eq!(out, &["excellent", "need work"]);
}

#[test]
fn test_enum_iteration_with_for() {
    let out = run_prints(
        r#"
        enum class Channel {
            RED, GREEN, BLUE
        }

        fun main() {
            var names = ""
            for (item in arrayOf(Channel.RED, Channel.GREEN, Channel.BLUE)) {
                names = names + item + ","
            }
            println(names)
        }
    "#,
    );
    // Real Kotlin agrees: string concatenation renders the NAMES.
    assert_eq!(out, &["RED,GREEN,BLUE,"]);
}

#[test]
fn test_enum_entry_order_with_when_subject() {
    let out = run_prints(
        r#"
        enum class Priority { FIRST, SECOND, THIRD }

        fun rank(priority: Priority): Int {
            return when (priority) {
                Priority.FIRST -> 1
                Priority.SECOND -> 2
                Priority.THIRD -> 3
            }
        }

        fun main() {
            println(rank(Priority.SECOND))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_enum_to_int_ordinal_like_behavior() {
    let out = run_prints(
        r#"
        enum class Dice {
            ONE, TWO, THREE
        }

        fun main() {
            val a = Dice.ONE
            val b = Dice.TWO
            val c = Dice.THREE
            println(a)
            println(c)
        }
    "#,
    );
    // Real Kotlin agrees: names, not ordinals.
    assert_eq!(out, &["ONE", "THREE"]);
}

#[test]
fn test_enum_payloads_and_references() {
    let out = run_prints(
        r#"
        enum class Planet(val mass: Int) {
            MERCURY(1),
            EARTH(2),
            MARS(3)
        }

        fun main() {
            val current = Planet.EARTH
            println(current.mass)
            println(Planet.MARS.mass)
        }
    "#,
    );
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_enum_when_fallback() {
    let out = run_prints(
        r#"
        enum class Mode { ON, OFF }

        fun describe(mode: Mode): String {
            return when (mode) {
                Mode.ON -> "enabled"
                Mode.OFF -> "disabled"
            }
        }

        fun main() {
            println(describe(Mode.ON))
            println(describe(Mode.OFF))
        }
    "#,
    );
    assert_eq!(out, &["enabled", "disabled"]);
}

#[test]
fn test_enum_reference_in_boolean() {
    let out = run_prints(
        r#"
        enum class Flag { TRUE, FALSE }

        fun main() {
            val f = Flag.TRUE
            println(f == Flag.TRUE)
            println(f != Flag.FALSE)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_enum_with_custom_payload_simple() {
    let out = run_prints(
        r#"
enum class Size(val value: Int) { SMALL(1), MEDIUM(2), LARGE(3) }; fun main() { println(Size.MEDIUM.value); println(Size.LARGE.value) }
"#,
    );
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_enum_boolean_like_values() {
    let out = run_prints(
        r#"
enum class Switch { ON, OFF }; fun main() { val s = Switch.OFF; println(s == Switch.OFF) }
"#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_enum_when_nested_conditions() {
    let out = run_prints(
        r#"
enum class Grade { A, B, C }; fun rank(g: Grade): Int { return when (g) { Grade.A -> 3; Grade.B -> 2; Grade.C -> 1 } }; fun main() { println(rank(Grade.B)); println(rank(Grade.A)) }
"#,
    );
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_enum_array_iteration_sum() {
    let out = run_prints(
        r#"
enum class Digit { D0, D1, D2, D3 }; fun main() { var n = 0; for (d in arrayOf(Digit.D0, Digit.D1, Digit.D2, Digit.D3)) { n += d.ordinal }; println(n) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_enum_compare_chain() {
    let out = run_prints(
        r#"
enum class Signal { RED, GREEN, YELLOW }; fun main() { val a = Signal.RED; val b = Signal.GREEN; println(a == b); println(a != b) }
"#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_enum_as_function_argument() {
    let out = run_prints(
        r#"
enum class Mode { FAST, SLOW }; fun step(mode: Mode): Int { return if (mode == Mode.FAST) 2 else 1 }; fun main() { println(step(Mode.FAST)); println(step(Mode.SLOW)) }
"#,
    );
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_enum_payload_and_property_access() {
    let out = run_prints(
        r#"
enum class Zone(val id: Int) { A(10), B(20), C(30) }; fun main() { println(Zone.B.id) }
"#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_enum_nested_when_without_default() {
    let out = run_prints(
        r#"
enum class Tag { ONE, TWO }; fun describe(t: Tag): String { return when (t) { Tag.ONE -> "first"; Tag.TWO -> "second" } }; fun main() { println(describe(Tag.ONE)); println(describe(Tag.TWO)) }
"#,
    );
    assert_eq!(out, &["first", "second"]);
}

#[test]
fn test_enum_multiple_values_and_sum() {
    let out = run_prints(
        r#"
enum class Piece { A, B, C, D }; fun main() { var total = 0; for (p in arrayOf(Piece.A, Piece.B, Piece.C, Piece.D)) { total += p.ordinal }; println(total) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_enum_three_state_machine() {
    let out = run_prints(
        r#"
enum class State { START, PROCESS, END }; fun status(s: State): Int { return when (s) { State.START -> 0; State.PROCESS -> 1; State.END -> 2 } }; fun main() { println(status(State.PROCESS)) }
"#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_enum_with_payload_chain_calc() {
    let out = run_prints(
        r#"
enum class Level(val factor: Int) { LOW(1), MID(2), HIGH(3) }; fun main() { val selected = Level.HIGH; println(selected.factor * 2) }
"#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_enum_aliasing_via_copy() {
    let out = run_prints(
        r#"
enum class Kind { A, B, C }; fun main() { val a = Kind.A; val b = a; println(a == b) }
"#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_enum_boolean_expression_match() {
    let out = run_prints(
        r#"
enum class Flag { YES, NO }; fun main() { val answer = Flag.YES; println(if (answer == Flag.YES) "ok" else "no") }
"#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_enum_long_chain_in_for() {
    let out = run_prints(
        r#"
enum class Step { ONE, TWO, THREE, FOUR }; fun main() { var hit = 0; for (s in arrayOf(Step.ONE, Step.TWO, Step.THREE, Step.FOUR)) { hit += 1 }; println(hit) }
"#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_enum_values_reflects_declaration_order() {
    let out = run_prints(
        r#"
        enum class Color { RED, GREEN, BLUE }

        fun main() {
            val values = Color.values()
            var names = ""
            for (value in values) {
                names += value.name + "|"
            }
            println(names)
            println(values.size)
            println(values[0] == Color.RED)
        }
    "#,
    );
    assert_eq!(out, &["RED|GREEN|BLUE|", "3", "true"]);
}

#[test]
fn test_enum_ordinal_and_name_contract() {
    let out = run_prints(
        r#"
        enum class Planet { MERCURY, EARTH, MARS }

        fun main() {
            val target = Planet.EARTH
            println(target.name)
            println(target.ordinal)
            println(Planet.MERCURY.ordinal < Planet.EARTH.ordinal)
        }
    "#,
    );
    assert_eq!(out, &["EARTH", "1", "true"]);
}

#[test]
fn test_enum_value_of_lookup() {
    let out = run_prints(
        r#"
        enum class State { START, RUN, STOP }

        fun main() {
            println(State.valueOf("RUN"))
            try {
                State.valueOf("PAUSE")
                println("found")
            } catch (e: IllegalArgumentException) {
                println("missing")
            }
        }
    "#,
    );
    assert_eq!(out, &["RUN", "missing"]);
}

#[test]
fn test_enum_with_abstract_member_dispatch() {
    let out = run_prints(
        r#"
        enum class Operation {
            ADD {
                override fun apply(a: Int, b: Int): Int = a + b
            },
            SUBTRACT {
                override fun apply(a: Int, b: Int): Int = a - b
            },
            MULTIPLY {
                override fun apply(a: Int, b: Int): Int = a * b
            };

            abstract fun apply(a: Int, b: Int): Int
        }

        fun main() {
            println(Operation.ADD.apply(4, 2))
            println(Operation.SUBTRACT.apply(7, 3))
            println(Operation.MULTIPLY.apply(3, 5))
        }
    "#,
    );
    assert_eq!(out, &["6", "4", "15"]);
}

#[test]
fn test_enum_set_membership_and_inclusion() {
    let out = run_prints(
        r#"
enum class Flag { ON, OFF, UNKNOWN }; fun main() { val active = Flag.ON; val allowed = setOf(Flag.ON, Flag.OFF); println(active in allowed); println(Flag.UNKNOWN in allowed) }
"#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_enum_entries_roundtrip_string_lookup() {
    let out = run_prints(
        r#"
        enum class Mode { READ, WRITE, EXECUTE }

        fun describe(value: String): String {
            return try {
                val mode = Mode.valueOf(value)
                "ok:" + mode.name
            } catch (e: Exception) {
                "bad"
            }
        }

        fun main() {
            println(describe("WRITE"))
            println(describe("bad"))
        }
    "#,
    );
    assert_eq!(out, &["ok:WRITE", "bad"]);
}

#[test]
fn test_enum_values_returns_fresh_array() {
    let out = run_prints(
        r#"
        enum class Color { RED, GREEN, BLUE }

        fun main() {
            val first = Color.values()
            val second = Color.values()
            first[0] = Color.GREEN
            println(second[0] == Color.RED)
            println(first[0] == Color.GREEN)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_enum_by_values_index_matches_ordinal() {
    let out = run_prints(
        r#"
        enum class Grade { A, B, C, D }

        fun main() {
            println(Grade.values()[0] == Grade.A)
            println(Grade.values()[2].ordinal)
            println(Grade.values().size)
        }
    "#,
    );
    assert_eq!(out, &["true", "2", "4"]);
}

#[test]
fn test_enum_nested_scope_reference() {
    let out = run_prints(
        r#"
        class Traffic {
            enum class Light { RED, YELLOW, GREEN }
        }

        fun main() {
            val current = Traffic.Light.GREEN
            println(current.name)
            println(current.ordinal)
        }
    "#,
    );
    assert_eq!(out, &["GREEN", "2"]);
}

#[test]
fn test_enum_used_in_map_keys() {
    let out = run_prints(
        r#"
        enum class Option { READ, WRITE, EXECUTE }

        fun main() {
            val permissions = mapOf(
                Option.READ to "read-only",
                Option.WRITE to "read-write",
                Option.EXECUTE to "admin"
            )
            println(permissions[Option.READ])
            println(permissions[Option.EXECUTE])
        }
    "#,
    );
    assert_eq!(out, &["read-only", "admin"]);
}

#[test]
fn test_enum_case_sensitive_value_of_lookup() {
    let out = run_prints(
        r#"
        enum class Flag { Enabled, Disabled }

        fun main() {
            try {
                println(Flag.valueOf("Enabled"))
            } catch (e: Exception) {
                println("bad")
            }

            try {
                println(Flag.valueOf("enabled"))
                println("should not happen")
            } catch (e: Exception) {
                println("missing")
            }
        }
    "#,
    );
    assert_eq!(out, &["Enabled", "missing"]);
}

#[test]
fn test_enum_when_expression_returns_defaulted_value() {
    let out = run_prints(
        r#"
        enum class Level { LOW, MEDIUM, HIGH }

        fun describe(level: Level): String {
            return when (level) {
                Level.LOW -> "low"
                Level.MEDIUM -> "medium"
                Level.HIGH -> "high"
            }
        }

        fun main() {
            val level = if (2 > 1) Level.HIGH else Level.LOW
            println(describe(level))
            println(describe(Level.MEDIUM))
        }
    "#,
    );
    assert_eq!(out, &["high", "medium"]);
}
