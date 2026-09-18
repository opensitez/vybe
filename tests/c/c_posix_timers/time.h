#ifndef _COMPAT_POSIX_TIMERS_TIME_H
#define _COMPAT_POSIX_TIMERS_TIME_H

#include_next <time.h>
#include <errno.h>
#include <string.h>

#ifndef TIMER_ABSTIME
#define TIMER_ABSTIME 1
#endif

struct itimerspec {
    struct timespec it_interval;
    struct timespec it_value;
};

typedef int timer_t;

struct __compat_timer {
    int active;
    struct itimerspec its;
};

static struct __compat_timer __timers[16];

struct sigevent;

static inline int timer_create(clockid_t clockid, struct sigevent *evp, timer_t *timerid) {
    (void)evp;
    if (clockid != CLOCK_REALTIME && clockid != CLOCK_MONOTONIC) {
        errno = EINVAL;
        return -1;
    }
    for (int i = 0; i < 16; i++) {
        if (!__timers[i].active) {
            __timers[i].active = 1;
            memset(&__timers[i].its, 0, sizeof(struct itimerspec));
            *timerid = i;
            return 0;
        }
    }
    errno = EAGAIN;
    return -1;
}

static inline int timer_delete(timer_t timerid) {
    if (timerid < 0 || timerid >= 16 || !__timers[timerid].active) {
        errno = EINVAL;
        return -1;
    }
    __timers[timerid].active = 0;
    return 0;
}

static inline int timer_settime(timer_t timerid, int flags, const struct itimerspec *new_value, struct itimerspec *old_value) {
    (void)flags;
    if (timerid < 0 || timerid >= 16 || !__timers[timerid].active) {
        errno = EINVAL;
        return -1;
    }
    if (old_value) {
        *old_value = __timers[timerid].its;
    }
    if (new_value) {
        __timers[timerid].its = *new_value;
    }
    return 0;
}

static inline int timer_gettime(timer_t timerid, struct itimerspec *curr_value) {
    if (timerid < 0 || timerid >= 16 || !__timers[timerid].active) {
        errno = EINVAL;
        return -1;
    }
    if (curr_value) {
        *curr_value = __timers[timerid].its;
    }
    return 0;
}

static inline int timer_getoverrun(timer_t timerid) {
    if (timerid < 0 || timerid >= 16 || !__timers[timerid].active) {
        errno = EINVAL;
        return -1;
    }
    return 0;
}

static inline int clock_nanosleep(clockid_t clock_id, int flags, const struct timespec *request, struct timespec *remain) {
    if (flags != 0 && flags != TIMER_ABSTIME) {
        return EINVAL;
    }
    if (!request) return EINVAL;
    if (request->tv_nsec < 0 || request->tv_nsec >= 1000000000L) return EINVAL;
    if (flags == TIMER_ABSTIME) {
        struct timespec now;
        clock_gettime(clock_id, &now);
        struct timespec diff;
        diff.tv_sec = request->tv_sec - now.tv_sec;
        diff.tv_nsec = request->tv_nsec - now.tv_nsec;
        if (diff.tv_nsec < 0) {
            diff.tv_sec--;
            diff.tv_nsec += 1000000000L;
        }
        if (diff.tv_sec < 0) {
            return 0;
        }
        return nanosleep(&diff, remain) == 0 ? 0 : errno;
    }
    return nanosleep(request, remain) == 0 ? 0 : errno;
}

#endif
