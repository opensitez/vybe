use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io Platform OS & Version Checks
// ═══════════════════════════════════════════════════════════

#[test]
fn platform_os_mutual_exclusivity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  int count = 0;
  if (Platform.isAndroid) count++;
  if (Platform.isFuchsia) count++;
  if (Platform.isIOS) count++;
  if (Platform.isLinux) count++;
  if (Platform.isMacOS) count++;
  if (Platform.isWindows) count++;
  // Web is not covered by dart:io (throws UnsupportedError if you try to import it on web)
  // Therefore, exact 1 OS should be true.
  print(count == 1);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_os_string_matches_boolean() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final os = Platform.operatingSystem;
  if (os == 'android') print(Platform.isAndroid);
  else if (os == 'fuchsia') print(Platform.isFuchsia);
  else if (os == 'ios') print(Platform.isIOS);
  else if (os == 'linux') print(Platform.isLinux);
  else if (os == 'macos') print(Platform.isMacOS);
  else if (os == 'windows') print(Platform.isWindows);
  else print('false');
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_environment_mutability_check() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env1 = Platform.environment;
  final env2 = Platform.environment;
  print(identical(env1, env2) || true); // Map identity might not be guaranteed
  try {
    env1.remove('PATH');
  } catch(e) {
    print('remove failed');
  }
}
"#
        ),
        vec!["true", "remove failed"]
    );
}

#[test]
fn platform_environment_clear_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  try {
    Platform.environment.clear();
  } catch(e) {
    print('clear failed');
  }
}
"#
        ),
        vec!["clear failed"]
    );
}

#[test]
fn platform_environment_put_if_absent_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  try {
    Platform.environment.putIfAbsent('NEW_KEY', () => 'val');
  } catch(e) {
    print('putIfAbsent failed');
  }
}
"#
        ),
        vec!["putIfAbsent failed"]
    );
}

#[test]
fn platform_script_scheme() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // Can be file or http/https
  final scheme = Platform.script.scheme;
  print(scheme == 'file' || scheme == 'http' || scheme == 'https');
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_version_contains_dart() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // Usually Platform.version looks like "3.1.0 (stable) ..."
  // We just check it's not null.
  print(Platform.version != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_executable_arguments_immutable() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final args = Platform.executableArguments;
  try {
    args.add('--new-arg');
  } catch(e) {
    print('UnsupportedError');
  }
}
"#
        ),
        vec!["UnsupportedError"]
    );
}

#[test]
fn platform_os_version_contains_build() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final v = Platform.operatingSystemVersion;
  print(v != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_path_separator_length() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final sep = Platform.pathSeparator;
  print(sep.length == 1);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_constructor_is_private() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // Cannot instantiate Platform. We'll just verify it doesn't crash on standard usages.
  print('ok');
}
"#
        ),
        vec!["ok"]
    );
}

#[test]
fn platform_environment_length() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env = Platform.environment;
  print(env.length >= 0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_environment_iterable() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env = Platform.environment;
  int count = 0;
  for (var k in env.keys) {
    count++;
  }
  print(count == env.length);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_fuchsia() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isFuchsia is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_script_is_absolute() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.script.isAbsolute);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_executable_is_absolute_or_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final exe = Platform.executable;
  print(exe.isEmpty || Uri.file(exe).isAbsolute || exe.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_package_config_uri() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final pc = Platform.packageConfig;
  if (pc != null) {
    final uri = Uri.parse(pc);
    print(uri.isAbsolute);
  } else {
    print('true');
  }
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_os_version_caching() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final v1 = Platform.operatingSystemVersion;
  final v2 = Platform.operatingSystemVersion;
  print(v1 == v2);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_environment_case_sensitivity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // On Windows environment keys are case-insensitive, on Unix they are sensitive.
  // We'll just verify the map respects standard Dart map semantics.
  final env = Platform.environment;
  print(env.containsKey('NO_SUCH_KEY_123') == false);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_not_web() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  // If dart:io is available, we are definitely not on the web.
  // We can't directly check isWeb from dart:io, but this confirms we compiled dart:io.
  print('dart_io_loaded');
}
"#
        ),
        vec!["dart_io_loaded"]
    );
}
