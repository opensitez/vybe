//! os/signal, os/user, syscall — breadth compile smokes.

use crate::helpers::*;

go_compile_cases! {
    // os/signal
    signal_notify => "package main; import \"os\"; import \"os/signal\"; import \"syscall\"; func main() { ch := make(chan os.Signal, 1); signal.Notify(ch, syscall.SIGINT) }",
    signal_stop => "package main; import \"os\"; import \"os/signal\"; import \"syscall\"; func main() { ch := make(chan os.Signal, 1); signal.Notify(ch, syscall.SIGTERM); signal.Stop(ch) }",
    signal_reset => "package main; import \"os/signal\"; import \"syscall\"; func main() { signal.Reset(syscall.SIGHUP) }",
    signal_notify_context => "package main; import \"context\"; import \"os/signal\"; import \"syscall\"; func main() { ctx, stop := signal.NotifyContext(context.Background(), syscall.SIGINT); defer stop(); _ = ctx }",
    signal_ignored => "package main; import \"os/signal\"; import \"syscall\"; func main() { _ = signal.Ignored(syscall.SIGPIPE) }",

    // os/user — lookup functions
    user_current => "package main; import \"os/user\"; func main() { _, _ = user.Current() }",
    user_lookup => "package main; import \"os/user\"; func main() { _, _ = user.Lookup(\"root\") }",
    user_lookup_id => "package main; import \"os/user\"; func main() { _, _ = user.LookupId(\"0\") }",
    user_lookup_group => "package main; import \"os/user\"; func main() { _, _ = user.LookupGroup(\"staff\") }",
    user_lookup_group_id => "package main; import \"os/user\"; func main() { _, _ = user.LookupGroupId(\"20\") }",

    // os/user — User fields
    user_username_field => "package main; import \"os/user\"; func main() { u, _ := user.Current(); if u != nil { _ = u.Username } }",
    user_uid_field => "package main; import \"os/user\"; func main() { u, _ := user.Current(); if u != nil { _ = u.Uid } }",
    user_gid_field => "package main; import \"os/user\"; func main() { u, _ := user.Current(); if u != nil { _ = u.Gid } }",
    user_name_field => "package main; import \"os/user\"; func main() { u, _ := user.Current(); if u != nil { _ = u.Name } }",
    user_home_dir_field => "package main; import \"os/user\"; func main() { u, _ := user.Current(); if u != nil { _ = u.HomeDir } }",

    // os/user — Group fields
    user_group_name_field => "package main; import \"os/user\"; func main() { g, _ := user.LookupGroup(\"staff\"); if g != nil { _ = g.Name } }",
    user_group_gid_field => "package main; import \"os/user\"; func main() { g, _ := user.LookupGroupId(\"20\"); if g != nil { _ = g.Gid } }",

    // syscall — process identity
    syscall_getpid => "package main; import \"syscall\"; func main() { _ = syscall.Getpid() }",
    syscall_getppid => "package main; import \"syscall\"; func main() { _ = syscall.Getppid() }",
    syscall_getuid => "package main; import \"syscall\"; func main() { _ = syscall.Getuid() }",
    syscall_getgid => "package main; import \"syscall\"; func main() { _ = syscall.Getgid() }",
    syscall_geteuid => "package main; import \"syscall\"; func main() { _ = syscall.Geteuid() }",
    syscall_getegid => "package main; import \"syscall\"; func main() { _ = syscall.Getegid() }",

    // syscall — file descriptors
    syscall_open => "package main; import \"syscall\"; func main() { fd, _ := syscall.Open(\".\", syscall.O_RDONLY, 0); if fd >= 0 { syscall.Close(fd) } }",
    syscall_close => "package main; import \"syscall\"; func main() { fd, _ := syscall.Open(\".\", syscall.O_RDONLY, 0); if fd >= 0 { _ = syscall.Close(fd) } }",
    syscall_read => "package main; import \"syscall\"; func main() { fd, _ := syscall.Open(\".\", syscall.O_RDONLY, 0); if fd >= 0 { defer syscall.Close(fd); buf := make([]byte, 8); _, _ = syscall.Read(fd, buf) } }",
    syscall_write => "package main; import \"syscall\"; func main() { _, _ = syscall.Write(syscall.Stdout, []byte(\"x\")) }",

    // syscall — stat family
    syscall_stat => "package main; import \"syscall\"; type StatT = syscall.Stat_t; func main() { var st StatT; _, _ = syscall.Stat(\".\", &st) }",
    syscall_fstat => "package main; import \"syscall\"; type StatT = syscall.Stat_t; func main() { var st StatT; _, _ = syscall.Fstat(syscall.Stdout, &st) }",
    syscall_lstat => "package main; import \"syscall\"; type StatT = syscall.Stat_t; func main() { var st StatT; _, _ = syscall.Lstat(\".\", &st) }",

    // syscall — directories and links
    syscall_mkdir => "package main; import \"syscall\"; func main() { _ = syscall.Mkdir(\"tmpdir\", 0755) }",
    syscall_rmdir => "package main; import \"syscall\"; func main() { _ = syscall.Rmdir(\"tmpdir\") }",
    syscall_unlink => "package main; import \"syscall\"; func main() { _ = syscall.Unlink(\"missing.txt\") }",
    syscall_rename => "package main; import \"syscall\"; func main() { _ = syscall.Rename(\"a.txt\", \"b.txt\") }",
    syscall_chdir => "package main; import \"syscall\"; func main() { _ = syscall.Chdir(\".\") }",
    syscall_getcwd => "package main; import \"syscall\"; func main() { _, _ = syscall.Getcwd() }",

    // syscall — sockets
    syscall_socket => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { syscall.Close(fd) } }",
    syscall_bind => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); sa := &syscall.SockaddrInet4{Port: 0}; _ = syscall.Bind(fd, sa) } }",
    syscall_listen => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); _ = syscall.Listen(fd, 1) } }",
    syscall_accept => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); _, _ = syscall.Accept(fd) } }",
    syscall_connect => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); sa := &syscall.SockaddrInet4{Port: 80}; _ = syscall.Connect(fd, sa) } }",
    syscall_setsockopt_int => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); _ = syscall.SetsockoptInt(fd, syscall.SOL_SOCKET, syscall.SO_REUSEADDR, 1) } }",
    syscall_getsockopt_int => "package main; import \"syscall\"; func main() { fd, _ := syscall.Socket(syscall.AF_INET, syscall.SOCK_STREAM, 0); if fd >= 0 { defer syscall.Close(fd); _, _ = syscall.GetsockoptInt(fd, syscall.SOL_SOCKET, syscall.SO_TYPE) } }",

    // syscall — pipes and dup
    syscall_pipe => "package main; import \"syscall\"; func main() { var p [2]int; _, _ = syscall.Pipe(p[:]) }",
    syscall_dup => "package main; import \"syscall\"; func main() { _, _ = syscall.Dup(syscall.Stdout) }",

    // syscall — process control
    syscall_kill => "package main; import \"syscall\"; func main() { _ = syscall.Kill(syscall.Getpid(), syscall.SIGCONT) }",
    syscall_wait4 => "package main; import \"syscall\"; type WaitStatus = syscall.WaitStatus; func main() { var status WaitStatus; _, _ = syscall.Wait4(-1, &status, syscall.WNOHANG, nil) }",

    // syscall — time and environment
    syscall_gettimeofday => "package main; import \"syscall\"; type Timeval = syscall.Timeval; func main() { var tv Timeval; _ = syscall.Gettimeofday(&tv) }",
    syscall_environ => "package main; import \"syscall\"; func main() { _ = syscall.Environ() }",
    syscall_setenv => "package main; import \"syscall\"; func main() { _ = syscall.Setenv(\"VYBE_KEY\", \"1\") }",
    syscall_getenv => "package main; import \"syscall\"; func main() { _ = syscall.Getenv(\"PATH\") }",
    syscall_unsetenv => "package main; import \"syscall\"; func main() { _ = syscall.Unsetenv(\"VYBE_KEY\") }",
    syscall_clearenv => "package main; import \"syscall\"; func main() { syscall.Clearenv() }",

    // syscall — fcntl and flock
    syscall_fcntl => "package main; import \"syscall\"; func main() { _, _ = syscall.FcntlInt(syscall.Stdout, syscall.F_GETFL, 0) }",
    syscall_flock => "package main; import \"syscall\"; func main() { fd, _ := syscall.Open(\".\", syscall.O_RDONLY, 0); if fd >= 0 { defer syscall.Close(fd); _ = syscall.Flock(fd, syscall.LOCK_EX) } }",

    // syscall — mmap and seek
    syscall_mmap => "package main; import \"syscall\"; func main() { _, _, _ = syscall.Mmap(-1, 0, 4096, syscall.PROT_READ, syscall.MAP_ANON) }",
    syscall_munmap => "package main; import \"syscall\"; func main() { b, _, _ := syscall.Mmap(-1, 0, 4096, syscall.PROT_READ, syscall.MAP_ANON); if len(b) > 0 { _ = syscall.Munmap(b) } }",
    syscall_truncate => "package main; import \"syscall\"; func main() { _ = syscall.Truncate(\"file.txt\", 0) }",
    syscall_lseek => "package main; import \"syscall\"; func main() { _, _ = syscall.Seek(syscall.Stdout, 0, 0) }",

    // syscall — string conversion and errors
    syscall_byte_slice_to_string => "package main; import \"syscall\"; func main() { _ = syscall.ByteSliceToString([]byte(\"go\")) }",
    syscall_string_to_byte_slice => "package main; import \"syscall\"; func main() { _ = syscall.StringToByteSlice(\"go\") }",
    syscall_errno => "package main; import \"syscall\"; type Errno = syscall.Errno; func main() { var e Errno; _ = e.Error(); _ = syscall.ENOENT }",

    // syscall — chmod, chown, links
    syscall_chmod => "package main; import \"syscall\"; func main() { _ = syscall.Chmod(\".\", 0755) }",
    syscall_chown => "package main; import \"syscall\"; func main() { _ = syscall.Chown(\".\", -1, -1) }",

    // syscall — raw calls and constants
    syscall_raw_syscall => "package main; import \"syscall\"; func main() { _, _, _ = syscall.RawSyscall(syscall.SYS_GETPID, 0, 0, 0) }",
    syscall_syscall => "package main; import \"syscall\"; func main() { _, _, _ = syscall.Syscall(syscall.SYS_GETUID, 0, 0, 0) }",
    syscall_fork_lock => "package main; import \"syscall\"; func main() { _ = syscall.ForkLock }",
    syscall_stdin_stdout_stderr => "package main; import \"syscall\"; func main() { _ = syscall.Stdin; _ = syscall.Stdout; _ = syscall.Stderr }",
}
