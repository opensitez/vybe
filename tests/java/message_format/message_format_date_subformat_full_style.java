// vybe-test: java/message_format/message_format_date_subformat_full_style
// origin: languages/java/tests/java/test_message_format.rs

public class Main {

    // A static String, NOT a StringBuilder. Calling a method on a bare static
    // FIELD receiver fails under Vybe with "undefined is not callable"
    // (measured): `SB.append(x)` throws while `StringBuilder l = SB;
    // l.append(x)` works, so the method is resolved from the receiver's
    // declared type at the call site and a static field carries none. String
    // concatenation onto a static field has no such problem.
    static String __buf = "";

    static void __p(Object o) {
        __buf = __buf + String.valueOf(o) + "\n";
    }

    static void __pr(Object o) {
        __buf = __buf + String.valueOf(o);
    }

    static void __check(String want) {
        String got = __buf;
        // The final `println` contributes a trailing newline that the expected
        // line vector never carried, so it is not part of the comparison.
        if (got.endsWith("\n")) {
            got = got.substring(0, got.length() - 1);
        }
        if (!got.equals(want)) {
            System.out.println("FAIL: want [" + want + "] got [" + got + "]");
            throw new RuntimeException("assertion failed");
        }
    }

    public static void main(String[] args) {
java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 0, 1, 0, 0, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("{0,date,full}", java.util.Locale.US); __p(mf.format(new Object[]{cal.getTime()}));
__check("Monday, January 1, 2024");
    }
}

