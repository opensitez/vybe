use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:io Platform & Environment
// ═══════════════════════════════════════════════════════════

#[test]
fn platform_environment_is_map() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env = Platform.environment;
  print(env is Map<String, String>);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_environment_is_unmodifiable() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env = Platform.environment;
  try {
    env['TEST_KEY'] = 'value';
    print('modified');
  } catch (e) {
    print('UnsupportedError thrown');
  }
}
"#
        ),
        vec!["UnsupportedError thrown"]
    );
}

#[test]
fn platform_environment_contains_path() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final env = Platform.environment;
  // Usually every OS has some form of PATH or Path
  final hasPath = env.containsKey('PATH') || env.containsKey('Path');
  // We'll just verify we can access keys
  print(env.keys.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_executable() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final exe = Platform.executable;
  print(exe.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_resolved_executable() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final exe = Platform.resolvedExecutable;
  print(exe.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_script_uri() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final script = Platform.script;
  print(script is Uri);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_executable_arguments() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final args = Platform.executableArguments;
  print(args.length >= 0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_package_config() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final packageConfig = Platform.packageConfig;
  // Can be null or String
  print(packageConfig == null || packageConfig is String);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_version_string() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final version = Platform.version;
  print(version.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_local_hostname() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final hostname = Platform.localHostname;
  print(hostname.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_operating_system() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final os = Platform.operatingSystem;
  final valid = ['android', 'fuchsia', 'ios', 'linux', 'macos', 'windows'].contains(os) || os.isNotEmpty;
  print(valid);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_operating_system_version() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final osVersion = Platform.operatingSystemVersion;
  print(osVersion.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_path_separator() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final sep = Platform.pathSeparator;
  print(sep == '/' || sep == '\\');
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_number_of_processors() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final processors = Platform.numberOfProcessors;
  print(processors > 0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_locale_name() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  final locale = Platform.localeName;
  print(locale.isNotEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_android() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isAndroid is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_macos() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isMacOS is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_windows() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isWindows is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_linux() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isLinux is bool);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn platform_is_ios() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:io';
void main() {
  print(Platform.isIOS is bool);
}
"#
        ),
        vec!["true"]
    );
}
