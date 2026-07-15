use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn posix_time_complex_formatting() {
    assert_eq!(
        run_c(
            r#"
#define _POSIX_C_SOURCE 200809L
#include <time.h>
#include <string.h>
#include <stdlib.h>

int main() {
    setenv("TZ", "UTC", 1);
    tzset();

    struct tm timeinfo = {0};
    timeinfo.tm_year = 2023 - 1900;
    timeinfo.tm_mon = 10 - 1; // October
    timeinfo.tm_mday = 15;
    timeinfo.tm_hour = 14;
    timeinfo.tm_min = 30;
    timeinfo.tm_sec = 45;
    timeinfo.tm_isdst = 0;
    
    time_t t = mktime(&timeinfo);
    
    char buf[128];
    // ISO 8601 format
    strftime(buf, sizeof(buf), "%Y-%m-%dT%H:%M:%SZ", gmtime(&t));
    printf("%s", buf);
    return 0;
}
    "#
        ),
        vec!["2023-10-15T14:30:45Z"]
    );
}

#[test]
fn posix_gettimeofday_nanosleep() {
    assert_eq!(
        run_c(
            r#"
#define _POSIX_C_SOURCE 200809L
#include <sys/time.h>
#include <time.h>

int main() {
    struct timeval tv1, tv2;
    gettimeofday(&tv1, NULL);
    
    struct timespec req = {0, 50000000}; // 50ms
    nanosleep(&req, NULL);
    
    gettimeofday(&tv2, NULL);
    
    long elapsed_us = (tv2.tv_sec - tv1.tv_sec) * 1000000L + (tv2.tv_usec - tv1.tv_usec);
    printf("%d", elapsed_us >= 40000); // Allow some OS scheduling tolerance
    return 0;
}
    "#
        ),
        vec!["1"]
    );
}
