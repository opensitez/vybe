// vybe-test: js/oop_patterns_advanced/template_method_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class Report {
    generate() {
        return [this.header(), this.body(), this.footer()].join("|");
    }
    header() { return "Report"; }
    footer() { return "End"; }
    body() { throw new Error("override body"); }
}
class SalesReport extends Report {
    body() { return "Sales: 1000"; }
}
class HRReport extends Report {
    header() { return "HR Report"; }
    body() { return "Staff: 50"; }
}
__check(__line(new SalesReport().generate()), "Report|Sales: 1000|End");
__check(__line(new HRReport().generate()), "HR Report|Staff: 50|End");
