// PHP date helpers — JS-source polyfills bundled as `__vybe_php_*`.
//
// Composes `new Date(...)` + the standard `getFullYear` / `getMonth` /
// `getDate` / etc. methods (resolved by the JS profile to ECMA host
// fns) so PHP gets `checkdate` / `getdate` without PHP-specific Rust.

// checkdate(month, day, year) — true iff the components form a real
// calendar date. JS rolls over invalid dates (Feb 30 → Mar 2), so we
// construct a Date and verify each component round-tripped.
function __vybe_php_checkdate(month, day, year) {
    if (month < 1 || month > 12) return false;
    if (day < 1 || day > 31) return false;
    if (year < 1 || year > 32767) return false;
    var d = new Date(year, month - 1, day);
    if (d.getFullYear() !== year) return false;
    if (d.getMonth() !== month - 1) return false;
    if (d.getDate() !== day) return false;
    return true;
}

// getdate(timestamp?) — returns an assoc array with date components.
// Defaults to current time. Mirrors PHP's keys exactly.
function __vybe_php_getdate(ts) {
    var d;
    if (ts === undefined || ts === null) {
        d = new Date();
    } else {
        d = new Date(ts * 1000);
    }
    var weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
    var month = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"];
    var info = {};
    info.seconds = d.getSeconds();
    info.minutes = d.getMinutes();
    info.hours   = d.getHours();
    info.mday    = d.getDate();
    info.wday    = d.getDay();
    info.mon     = d.getMonth() + 1;
    info.year    = d.getFullYear();
    info.weekday = weekday[d.getDay()];
    info.month   = month[d.getMonth()];
    return info;
}
