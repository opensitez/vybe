// vybe-test: js/regex_comprehensive/regex_match_all_groups
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const html = '<a href="http://foo.com">Foo</a> <a href="http://bar.com">Bar</a>';
const re = /<a href="([^"]+)">([^<]+)<\/a>/g;
const links = [...html.matchAll(re)].map(m => `${m[2]}:${m[1]}`);
console.log(links.join("|"));
