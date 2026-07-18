use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(enum_basic, "System.out.println(Color.RED);", "enum Color { RED, GREEN, BLUE }", "RED");
jm!(enum_ordinal, "System.out.println(Color.GREEN.ordinal());", "enum Color { RED, GREEN, BLUE }", "1");
jm!(enum_name, "System.out.println(Color.BLUE.name());", "enum Color { RED, GREEN, BLUE }", "BLUE");
jm!(enum_switch, "System.out.println(sw(Color.YELLOW));", "enum Color { RED, GREEN, YELLOW, BLUE } int sw(Color c) { switch(c){case RED:return 1;case GREEN:return 2;case YELLOW:return 3;default:return 0;} }", "3");
jm!(enum_switch_return, "System.out.println(sw(Color.RED));", "enum Color { RED, GREEN, BLUE } static int sw(Color c) { switch(c){case RED:return 10; case GREEN:return 20; default:return 30;} }", "10");
jm!(enum_with_fields, "System.out.println(Level.HIGH.value);", "enum Level { LOW(1), MEDIUM(2), HIGH(3); int value; Level(int v) { value = v; } }", "3");
jm!(enum_with_method, "System.out.println(Mode.A.next());", "enum Mode { A, B; Mode next() { return values()[1]; } }", "B");
jm!(enum_in_array, "System.out.println(Mode2.values().length);", "enum Mode2 { A, B, C }", "3");
jm!(enum_to_string, "System.out.println(Mode3.B.toString());", "enum Mode3 { A, B, C }", "B");
jm!(enum_compare, "System.out.println(Mode4.A == Mode4.A);", "enum Mode4 { A, B }", "true");
jm!(enum_compare_values, "System.out.println(Mode5.A == Mode5.B);", "enum Mode5 { A, B }", "false");
jm!(enum_hash_like, "System.out.println(Mode6.A.toString().equals(\"A\"));", "enum Mode6 { A, B }", "true");
jm!(enum_with_code, "System.out.println(Status.ERROR.code);", "enum Status { OK(0), ERROR(1); int code; Status(int code) { this.code = code; } }", "1");
jm!(enum_sum_code, "System.out.println(sumCodes());", "enum Status2 { A(1), B(2), C(3); int code; Status2(int code) { this.code = code; } static int sumCodes() { int s=0; for(Status2 s2: values()) s += s2.code; return s; } static int sumCodes(){ return sumCodes(); } }", "6");
jm!(enum_in_switch_default, "System.out.println(label(ColorX.UNKNOWN));", "enum ColorX { RED, GREEN, UNKNOWN } static String label(ColorX c) { switch(c){case RED:return \"r\";case GREEN:return \"g\";default:return \"x\";} }", "x");
jm!(enum_of_ordinal_access, "System.out.println(Step.values()[0]);", "enum Step { START, MIDDLE, END }", "START");
jm!(enum_in_ctor, "System.out.println(Power.on(Mode8.HIGH));", "enum Mode8 { LOW(1), HIGH(2); int value; Mode8(int value) { this.value = value; } static int on(Mode8 m) { return m.value; } }", "2");
jm!(enum_nested_call, "System.out.println(SystemType.Database.label());", "enum SystemType { Database { String label() { return \"db\"; } }, Cache { String label() { return \"cache\"; } }; String label() { return \"\"; } }", "db");
jm!(enum_ternary, "System.out.println(Priority.LOW.code + "," + Priority.HIGH.code);", "enum Priority { LOW(1), HIGH(10); int code; Priority(int code){ this.code = code; } }", "1,10");
jm!(enum_method_dispatch, "System.out.println(Pipe.YES.isEnabled());", "enum Pipe { YES, NO; boolean isEnabled(){ return this == YES; } }", "true");
jm!(enum_constructor_chain, "System.out.println(Phase.FIRST.value);", "enum Phase { FIRST(1), SECOND(2); int value; Phase(int v){ value=v; } }", "1");
jm!(enum_with_static_lookup, "System.out.println(fromCode(2));", "enum Tag { A(1), B(2); int code; Tag(int c){ code = c; } static Tag fromCode(int c) { for (Tag t: values()) if (t.code == c) return t; return null; } }", "B");
jm!(enum_boolean_check, "System.out.println(Status2.OK.ok());", "enum Status2 { OK, FAIL; boolean ok() { return this == OK; } }", "true");
jm!(enum_length_string, "System.out.println(Mode9.values().length);", "enum Mode9 { A, B, C, D }", "4");
jm!(enum_plus_name, "System.out.println(Trait.A + "," + Trait.B);", "enum Trait { A, B }", "A,B");
jm!(enum_in_condition, "System.out.println(Trait2.A.equals(Trait2.A) ? 1 : 0);", "enum Trait2 { A, B }", "1");
jm!(enum_as_return, "System.out.println(current());", "enum Kind { A, B } static Kind current() { return Kind.B; }", "B");
jm!(enum_casting_to_int, "System.out.println(Kind2.B.ordinal() + 1);", "enum Kind2 { A, B, C }", "2");
