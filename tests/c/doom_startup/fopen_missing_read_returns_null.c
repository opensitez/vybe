#include <stdio.h>

int main(void)
{
    FILE *f = fopen("definitely_missing_vybe_doom_config.cfg", "r");
    return f == 0 ? 0 : 1;
}
