#include <stdio.h>
int __vybe_c_charptr_visible_0;
int visits;
char *next_text(void) { visits++; return "abc"; }
int main(void) {
    int __vybe_c_charptr_visible_1 = 7;
    char output[4];
    snprintf(output, sizeof(output), next_text());
    snprintf(output, sizeof(output), next_text());
    snprintf(output, sizeof(output), next_text());
    return visits == 3 && __vybe_c_charptr_visible_1 == 7
        && output[0] == 'a' && output[2] == 'c' && output[3] == 0 ? 0 : 1;
}
