typedef struct {
    int partime;
} wbstartstruct_t;

wbstartstruct_t wminfo;

int main(void) {
    unsigned char bytes[4] = {11, 0, 0, 0};
    int flag = 1;
    if (flag) {
        int cpars32 = 0;
        memcpy(&cpars32, bytes, sizeof(int));
        wminfo.partime = 35 * cpars32;
    } else {
        wminfo.partime = 0;
    }
    return wminfo.partime == 385 ? 0 : 1;
}
