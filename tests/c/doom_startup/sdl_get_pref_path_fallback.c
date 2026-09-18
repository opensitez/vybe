extern char *SDL_GetPrefPath(const char *org, const char *app);
int strlen(const char *s);

int main(void)
{
    char *path = SDL_GetPrefPath("", "chocolate-doom");
    if (path == 0)
    {
        return 1;
    }
    return path[strlen(path) - 1] == '/' ? 0 : 1;
}
