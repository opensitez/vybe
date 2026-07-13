use super::helpers::run_c;

#[test]
fn test_regex_probe() {
    assert_eq!(
        run_c("
#include <stdio.h>
#include <regex.h>

int main() {
    regex_t regex;
    int ret = regcomp(&regex, \"^a[b-d]e$\", REG_EXTENDED);
    if (ret) {
        printf(\"compile failed\\n\");
        return 1;
    }
    ret = regexec(&regex, \"ace\", 0, NULL, 0);
    if (!ret) {
        printf(\"match\\n\");
    } else {
        printf(\"no match\\n\");
    }
    regfree(&regex);
    return 0;
}
        "),
        vec!["match"]
    );
}
