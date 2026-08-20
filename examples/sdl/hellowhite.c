#include <stdint.h>
  typedef uint32_t Uint32; typedef int32_t Sint32;
  typedef void *SDL_Window; typedef void *SDL_Surface;
  typedef struct SDL_Rect { Sint32 x, y, w, h; } SDL_Rect;
  extern int SDL_Init(Uint32 f);
  extern void SDL_Quit(void);
  extern SDL_Window *SDL_CreateWindow(const char*, Sint32, Sint32, Sint32, Sint32, Uint32);
  extern SDL_Surface *SDL_GetWindowSurface(SDL_Window*);
  extern int SDL_FillRect(SDL_Surface*, const SDL_Rect*, Uint32);
  extern int SDL_UpdateWindowSurface(SDL_Window*);
  extern Uint32 SDL_MapRGB(Uint32, Uint32, Uint32, Uint32);
  extern int SDL_DrawText(SDL_Surface*, const char*, Sint32, Sint32, Uint32);
  extern int SDL_ShowWindow(SDL_Window*);
  int main(void) {
      SDL_Init(32);
      SDL_Window *w = SDL_CreateWindow("TextTest", 100, 100, 400, 200, 4);
      SDL_Surface *s = SDL_GetWindowSurface(w);
      SDL_ShowWindow(w);
      SDL_Rect bg = {0, 0, 400, 200};
      SDL_FillRect(s, &bg, SDL_MapRGB(0, 255, 255, 255));   /* white */
      SDL_DrawText(s, "BLACK TEXT ON WHITE", 20, 40, SDL_MapRGB(0, 0, 0, 0));
      SDL_DrawText(s, "RED TEXT", 20, 90, SDL_MapRGB(0, 255, 0, 0));
      SDL_UpdateWindowSurface(w);
      SDL_Quit();
      return 0;
  }

