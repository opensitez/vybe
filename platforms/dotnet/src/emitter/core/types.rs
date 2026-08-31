use std::sync::LazyLock;

use super::super::types::{KnownTypeMapping, KnownTypeTarget};
use vybe_runtime::component_model::ConstructorTarget;

const KNOWN_CONSTANTS: &[&str] = &[
    "pi",
    "e",
    "maxvalue",
    "minvalue",
    "positiveinfinity",
    "negativeinfinity",
    "nan",
    "epsilon",
    "empty",
    "newline",
    "true",
    "false",
    "completedtask",
];

// π, e and τ are NOT part of this platform's surface — they are numbers, and
// `vybe_compiler::primitives::math` owns them for every language. The twelve
// rows that used to sit at the top of this table (`math.pi`, `Math.PI`,
// `system.math.pi`, `System.Math.PI`, ×3 concepts) made the dotnet platform the
// de-facto owner of `Math.PI`, so any language that wanted π had to reach a
// platform it has nothing to do with. `math::dotted_constant` matches the owner
// segment case-insensitively and so answers all four spellings on its own.
//
// What REMAINS here is genuinely .NET's: type limits spelled by .NET type names
// (`int.MaxValue`), and the framework enum ordinals (`CommandType`,
// `ConnectionState`, `RegexOptions`, `MsgBoxStyle`).
const NAMESPACE_CONSTANTS: &[(&str, f64)] = &[
    ("int.MaxValue", 2_147_483_647.0),
    ("int.MinValue", -2_147_483_648.0),
    ("double.MaxValue", f64::MAX),
    ("double.MinValue", -f64::MAX),
    ("double.NaN", f64::NAN),
    ("double.PositiveInfinity", f64::INFINITY),
    ("double.NegativeInfinity", f64::NEG_INFINITY),
    ("float.MaxValue", 3.4028235e38),
    ("float.MinValue", -3.4028235e38),
    // ⛔ `Epsilon` is the smallest SUBNORMAL, not `f64::MIN_POSITIVE` (the
    // smallest normal). .NET documents 4.94065645841247E-324 for Double and
    // 1.401298E-45 for Single, and the two differ — one shared value answered
    // both wrongly.
    ("double.Epsilon", 5e-324),
    ("float.Epsilon", 1.401298464324817e-45),
    ("single.Epsilon", 1.401298464324817e-45),
    ("single.NaN", f64::NAN),
    ("single.PositiveInfinity", f64::INFINITY),
    ("single.NegativeInfinity", f64::NEG_INFINITY),
    ("single.MaxValue", 3.4028235e38),
    ("single.MinValue", -3.4028235e38),
    ("char.MaxValue", 65535.0),
    ("char.MinValue", 0.0),
    // `System.Net.HttpStatusCode` — the status ordinals. A framework enum is a
    // NUMBER on this platform (`RegexOptions`, `CommandType`, `NumberStyles`
    // above all are), which is what makes `(int)status` answer: the shared cast
    // reads an `Int` role only for a DECLARED enum, and a platform type has no
    // `EnumDecl`. ⛔ The consequence is that `status.ToString()` renders `200`
    // where .NET renders `OK` — the same divergence every other framework enum
    // here already has, not a new one.
    //
    // ⛔ The lookup is a VERBATIM key match, with no suffix folding: measured,
    // `System.Globalization.NumberStyles.Integer` answers `NaN` while the bare
    // `NumberStyles.Integer` answers 7, because only the short key is listed.
    // Every row therefore appears four times — short and `System.Net.`-
    // qualified, each cased (C#) and lowercase (VB, whose profile lowercases
    // the key before looking it up).
    ("HttpStatusCode.Continue", 100.0),
    ("httpstatuscode.continue", 100.0),
    ("System.Net.HttpStatusCode.Continue", 100.0),
    ("system.net.httpstatuscode.continue", 100.0),
    ("HttpStatusCode.SwitchingProtocols", 101.0),
    ("httpstatuscode.switchingprotocols", 101.0),
    ("System.Net.HttpStatusCode.SwitchingProtocols", 101.0),
    ("system.net.httpstatuscode.switchingprotocols", 101.0),
    ("HttpStatusCode.Processing", 102.0),
    ("httpstatuscode.processing", 102.0),
    ("System.Net.HttpStatusCode.Processing", 102.0),
    ("system.net.httpstatuscode.processing", 102.0),
    ("HttpStatusCode.EarlyHints", 103.0),
    ("httpstatuscode.earlyhints", 103.0),
    ("System.Net.HttpStatusCode.EarlyHints", 103.0),
    ("system.net.httpstatuscode.earlyhints", 103.0),
    ("HttpStatusCode.OK", 200.0),
    ("httpstatuscode.ok", 200.0),
    ("System.Net.HttpStatusCode.OK", 200.0),
    ("system.net.httpstatuscode.ok", 200.0),
    ("HttpStatusCode.Created", 201.0),
    ("httpstatuscode.created", 201.0),
    ("System.Net.HttpStatusCode.Created", 201.0),
    ("system.net.httpstatuscode.created", 201.0),
    ("HttpStatusCode.Accepted", 202.0),
    ("httpstatuscode.accepted", 202.0),
    ("System.Net.HttpStatusCode.Accepted", 202.0),
    ("system.net.httpstatuscode.accepted", 202.0),
    ("HttpStatusCode.NonAuthoritativeInformation", 203.0),
    ("httpstatuscode.nonauthoritativeinformation", 203.0),
    ("System.Net.HttpStatusCode.NonAuthoritativeInformation", 203.0),
    ("system.net.httpstatuscode.nonauthoritativeinformation", 203.0),
    ("HttpStatusCode.NoContent", 204.0),
    ("httpstatuscode.nocontent", 204.0),
    ("System.Net.HttpStatusCode.NoContent", 204.0),
    ("system.net.httpstatuscode.nocontent", 204.0),
    ("HttpStatusCode.ResetContent", 205.0),
    ("httpstatuscode.resetcontent", 205.0),
    ("System.Net.HttpStatusCode.ResetContent", 205.0),
    ("system.net.httpstatuscode.resetcontent", 205.0),
    ("HttpStatusCode.PartialContent", 206.0),
    ("httpstatuscode.partialcontent", 206.0),
    ("System.Net.HttpStatusCode.PartialContent", 206.0),
    ("system.net.httpstatuscode.partialcontent", 206.0),
    ("HttpStatusCode.MultiStatus", 207.0),
    ("httpstatuscode.multistatus", 207.0),
    ("System.Net.HttpStatusCode.MultiStatus", 207.0),
    ("system.net.httpstatuscode.multistatus", 207.0),
    ("HttpStatusCode.AlreadyReported", 208.0),
    ("httpstatuscode.alreadyreported", 208.0),
    ("System.Net.HttpStatusCode.AlreadyReported", 208.0),
    ("system.net.httpstatuscode.alreadyreported", 208.0),
    ("HttpStatusCode.IMUsed", 226.0),
    ("httpstatuscode.imused", 226.0),
    ("System.Net.HttpStatusCode.IMUsed", 226.0),
    ("system.net.httpstatuscode.imused", 226.0),
    ("HttpStatusCode.MultipleChoices", 300.0),
    ("httpstatuscode.multiplechoices", 300.0),
    ("System.Net.HttpStatusCode.MultipleChoices", 300.0),
    ("system.net.httpstatuscode.multiplechoices", 300.0),
    ("HttpStatusCode.Ambiguous", 300.0),
    ("httpstatuscode.ambiguous", 300.0),
    ("System.Net.HttpStatusCode.Ambiguous", 300.0),
    ("system.net.httpstatuscode.ambiguous", 300.0),
    ("HttpStatusCode.MovedPermanently", 301.0),
    ("httpstatuscode.movedpermanently", 301.0),
    ("System.Net.HttpStatusCode.MovedPermanently", 301.0),
    ("system.net.httpstatuscode.movedpermanently", 301.0),
    ("HttpStatusCode.Moved", 301.0),
    ("httpstatuscode.moved", 301.0),
    ("System.Net.HttpStatusCode.Moved", 301.0),
    ("system.net.httpstatuscode.moved", 301.0),
    ("HttpStatusCode.Found", 302.0),
    ("httpstatuscode.found", 302.0),
    ("System.Net.HttpStatusCode.Found", 302.0),
    ("system.net.httpstatuscode.found", 302.0),
    ("HttpStatusCode.Redirect", 302.0),
    ("httpstatuscode.redirect", 302.0),
    ("System.Net.HttpStatusCode.Redirect", 302.0),
    ("system.net.httpstatuscode.redirect", 302.0),
    ("HttpStatusCode.SeeOther", 303.0),
    ("httpstatuscode.seeother", 303.0),
    ("System.Net.HttpStatusCode.SeeOther", 303.0),
    ("system.net.httpstatuscode.seeother", 303.0),
    ("HttpStatusCode.RedirectMethod", 303.0),
    ("httpstatuscode.redirectmethod", 303.0),
    ("System.Net.HttpStatusCode.RedirectMethod", 303.0),
    ("system.net.httpstatuscode.redirectmethod", 303.0),
    ("HttpStatusCode.NotModified", 304.0),
    ("httpstatuscode.notmodified", 304.0),
    ("System.Net.HttpStatusCode.NotModified", 304.0),
    ("system.net.httpstatuscode.notmodified", 304.0),
    ("HttpStatusCode.UseProxy", 305.0),
    ("httpstatuscode.useproxy", 305.0),
    ("System.Net.HttpStatusCode.UseProxy", 305.0),
    ("system.net.httpstatuscode.useproxy", 305.0),
    ("HttpStatusCode.Unused", 306.0),
    ("httpstatuscode.unused", 306.0),
    ("System.Net.HttpStatusCode.Unused", 306.0),
    ("system.net.httpstatuscode.unused", 306.0),
    ("HttpStatusCode.TemporaryRedirect", 307.0),
    ("httpstatuscode.temporaryredirect", 307.0),
    ("System.Net.HttpStatusCode.TemporaryRedirect", 307.0),
    ("system.net.httpstatuscode.temporaryredirect", 307.0),
    ("HttpStatusCode.RedirectKeepVerb", 307.0),
    ("httpstatuscode.redirectkeepverb", 307.0),
    ("System.Net.HttpStatusCode.RedirectKeepVerb", 307.0),
    ("system.net.httpstatuscode.redirectkeepverb", 307.0),
    ("HttpStatusCode.PermanentRedirect", 308.0),
    ("httpstatuscode.permanentredirect", 308.0),
    ("System.Net.HttpStatusCode.PermanentRedirect", 308.0),
    ("system.net.httpstatuscode.permanentredirect", 308.0),
    ("HttpStatusCode.BadRequest", 400.0),
    ("httpstatuscode.badrequest", 400.0),
    ("System.Net.HttpStatusCode.BadRequest", 400.0),
    ("system.net.httpstatuscode.badrequest", 400.0),
    ("HttpStatusCode.Unauthorized", 401.0),
    ("httpstatuscode.unauthorized", 401.0),
    ("System.Net.HttpStatusCode.Unauthorized", 401.0),
    ("system.net.httpstatuscode.unauthorized", 401.0),
    ("HttpStatusCode.PaymentRequired", 402.0),
    ("httpstatuscode.paymentrequired", 402.0),
    ("System.Net.HttpStatusCode.PaymentRequired", 402.0),
    ("system.net.httpstatuscode.paymentrequired", 402.0),
    ("HttpStatusCode.Forbidden", 403.0),
    ("httpstatuscode.forbidden", 403.0),
    ("System.Net.HttpStatusCode.Forbidden", 403.0),
    ("system.net.httpstatuscode.forbidden", 403.0),
    ("HttpStatusCode.NotFound", 404.0),
    ("httpstatuscode.notfound", 404.0),
    ("System.Net.HttpStatusCode.NotFound", 404.0),
    ("system.net.httpstatuscode.notfound", 404.0),
    ("HttpStatusCode.MethodNotAllowed", 405.0),
    ("httpstatuscode.methodnotallowed", 405.0),
    ("System.Net.HttpStatusCode.MethodNotAllowed", 405.0),
    ("system.net.httpstatuscode.methodnotallowed", 405.0),
    ("HttpStatusCode.NotAcceptable", 406.0),
    ("httpstatuscode.notacceptable", 406.0),
    ("System.Net.HttpStatusCode.NotAcceptable", 406.0),
    ("system.net.httpstatuscode.notacceptable", 406.0),
    ("HttpStatusCode.ProxyAuthenticationRequired", 407.0),
    ("httpstatuscode.proxyauthenticationrequired", 407.0),
    ("System.Net.HttpStatusCode.ProxyAuthenticationRequired", 407.0),
    ("system.net.httpstatuscode.proxyauthenticationrequired", 407.0),
    ("HttpStatusCode.RequestTimeout", 408.0),
    ("httpstatuscode.requesttimeout", 408.0),
    ("System.Net.HttpStatusCode.RequestTimeout", 408.0),
    ("system.net.httpstatuscode.requesttimeout", 408.0),
    ("HttpStatusCode.Conflict", 409.0),
    ("httpstatuscode.conflict", 409.0),
    ("System.Net.HttpStatusCode.Conflict", 409.0),
    ("system.net.httpstatuscode.conflict", 409.0),
    ("HttpStatusCode.Gone", 410.0),
    ("httpstatuscode.gone", 410.0),
    ("System.Net.HttpStatusCode.Gone", 410.0),
    ("system.net.httpstatuscode.gone", 410.0),
    ("HttpStatusCode.LengthRequired", 411.0),
    ("httpstatuscode.lengthrequired", 411.0),
    ("System.Net.HttpStatusCode.LengthRequired", 411.0),
    ("system.net.httpstatuscode.lengthrequired", 411.0),
    ("HttpStatusCode.PreconditionFailed", 412.0),
    ("httpstatuscode.preconditionfailed", 412.0),
    ("System.Net.HttpStatusCode.PreconditionFailed", 412.0),
    ("system.net.httpstatuscode.preconditionfailed", 412.0),
    ("HttpStatusCode.RequestEntityTooLarge", 413.0),
    ("httpstatuscode.requestentitytoolarge", 413.0),
    ("System.Net.HttpStatusCode.RequestEntityTooLarge", 413.0),
    ("system.net.httpstatuscode.requestentitytoolarge", 413.0),
    ("HttpStatusCode.RequestUriTooLong", 414.0),
    ("httpstatuscode.requesturitoolong", 414.0),
    ("System.Net.HttpStatusCode.RequestUriTooLong", 414.0),
    ("system.net.httpstatuscode.requesturitoolong", 414.0),
    ("HttpStatusCode.UnsupportedMediaType", 415.0),
    ("httpstatuscode.unsupportedmediatype", 415.0),
    ("System.Net.HttpStatusCode.UnsupportedMediaType", 415.0),
    ("system.net.httpstatuscode.unsupportedmediatype", 415.0),
    ("HttpStatusCode.RequestedRangeNotSatisfiable", 416.0),
    ("httpstatuscode.requestedrangenotsatisfiable", 416.0),
    ("System.Net.HttpStatusCode.RequestedRangeNotSatisfiable", 416.0),
    ("system.net.httpstatuscode.requestedrangenotsatisfiable", 416.0),
    ("HttpStatusCode.ExpectationFailed", 417.0),
    ("httpstatuscode.expectationfailed", 417.0),
    ("System.Net.HttpStatusCode.ExpectationFailed", 417.0),
    ("system.net.httpstatuscode.expectationfailed", 417.0),
    ("HttpStatusCode.MisdirectedRequest", 421.0),
    ("httpstatuscode.misdirectedrequest", 421.0),
    ("System.Net.HttpStatusCode.MisdirectedRequest", 421.0),
    ("system.net.httpstatuscode.misdirectedrequest", 421.0),
    ("HttpStatusCode.UnprocessableEntity", 422.0),
    ("httpstatuscode.unprocessableentity", 422.0),
    ("System.Net.HttpStatusCode.UnprocessableEntity", 422.0),
    ("system.net.httpstatuscode.unprocessableentity", 422.0),
    ("HttpStatusCode.UnprocessableContent", 422.0),
    ("httpstatuscode.unprocessablecontent", 422.0),
    ("System.Net.HttpStatusCode.UnprocessableContent", 422.0),
    ("system.net.httpstatuscode.unprocessablecontent", 422.0),
    ("HttpStatusCode.Locked", 423.0),
    ("httpstatuscode.locked", 423.0),
    ("System.Net.HttpStatusCode.Locked", 423.0),
    ("system.net.httpstatuscode.locked", 423.0),
    ("HttpStatusCode.FailedDependency", 424.0),
    ("httpstatuscode.faileddependency", 424.0),
    ("System.Net.HttpStatusCode.FailedDependency", 424.0),
    ("system.net.httpstatuscode.faileddependency", 424.0),
    ("HttpStatusCode.UpgradeRequired", 426.0),
    ("httpstatuscode.upgraderequired", 426.0),
    ("System.Net.HttpStatusCode.UpgradeRequired", 426.0),
    ("system.net.httpstatuscode.upgraderequired", 426.0),
    ("HttpStatusCode.PreconditionRequired", 428.0),
    ("httpstatuscode.preconditionrequired", 428.0),
    ("System.Net.HttpStatusCode.PreconditionRequired", 428.0),
    ("system.net.httpstatuscode.preconditionrequired", 428.0),
    ("HttpStatusCode.TooManyRequests", 429.0),
    ("httpstatuscode.toomanyrequests", 429.0),
    ("System.Net.HttpStatusCode.TooManyRequests", 429.0),
    ("system.net.httpstatuscode.toomanyrequests", 429.0),
    ("HttpStatusCode.RequestHeaderFieldsTooLarge", 431.0),
    ("httpstatuscode.requestheaderfieldstoolarge", 431.0),
    ("System.Net.HttpStatusCode.RequestHeaderFieldsTooLarge", 431.0),
    ("system.net.httpstatuscode.requestheaderfieldstoolarge", 431.0),
    ("HttpStatusCode.UnavailableForLegalReasons", 451.0),
    ("httpstatuscode.unavailableforlegalreasons", 451.0),
    ("System.Net.HttpStatusCode.UnavailableForLegalReasons", 451.0),
    ("system.net.httpstatuscode.unavailableforlegalreasons", 451.0),
    ("HttpStatusCode.InternalServerError", 500.0),
    ("httpstatuscode.internalservererror", 500.0),
    ("System.Net.HttpStatusCode.InternalServerError", 500.0),
    ("system.net.httpstatuscode.internalservererror", 500.0),
    ("HttpStatusCode.NotImplemented", 501.0),
    ("httpstatuscode.notimplemented", 501.0),
    ("System.Net.HttpStatusCode.NotImplemented", 501.0),
    ("system.net.httpstatuscode.notimplemented", 501.0),
    ("HttpStatusCode.BadGateway", 502.0),
    ("httpstatuscode.badgateway", 502.0),
    ("System.Net.HttpStatusCode.BadGateway", 502.0),
    ("system.net.httpstatuscode.badgateway", 502.0),
    ("HttpStatusCode.ServiceUnavailable", 503.0),
    ("httpstatuscode.serviceunavailable", 503.0),
    ("System.Net.HttpStatusCode.ServiceUnavailable", 503.0),
    ("system.net.httpstatuscode.serviceunavailable", 503.0),
    ("HttpStatusCode.GatewayTimeout", 504.0),
    ("httpstatuscode.gatewaytimeout", 504.0),
    ("System.Net.HttpStatusCode.GatewayTimeout", 504.0),
    ("system.net.httpstatuscode.gatewaytimeout", 504.0),
    ("HttpStatusCode.HttpVersionNotSupported", 505.0),
    ("httpstatuscode.httpversionnotsupported", 505.0),
    ("System.Net.HttpStatusCode.HttpVersionNotSupported", 505.0),
    ("system.net.httpstatuscode.httpversionnotsupported", 505.0),
    ("HttpStatusCode.VariantAlsoNegotiates", 506.0),
    ("httpstatuscode.variantalsonegotiates", 506.0),
    ("System.Net.HttpStatusCode.VariantAlsoNegotiates", 506.0),
    ("system.net.httpstatuscode.variantalsonegotiates", 506.0),
    ("HttpStatusCode.InsufficientStorage", 507.0),
    ("httpstatuscode.insufficientstorage", 507.0),
    ("System.Net.HttpStatusCode.InsufficientStorage", 507.0),
    ("system.net.httpstatuscode.insufficientstorage", 507.0),
    ("HttpStatusCode.LoopDetected", 508.0),
    ("httpstatuscode.loopdetected", 508.0),
    ("System.Net.HttpStatusCode.LoopDetected", 508.0),
    ("system.net.httpstatuscode.loopdetected", 508.0),
    ("HttpStatusCode.NotExtended", 510.0),
    ("httpstatuscode.notextended", 510.0),
    ("System.Net.HttpStatusCode.NotExtended", 510.0),
    ("system.net.httpstatuscode.notextended", 510.0),
    ("HttpStatusCode.NetworkAuthenticationRequired", 511.0),
    ("httpstatuscode.networkauthenticationrequired", 511.0),
    ("System.Net.HttpStatusCode.NetworkAuthenticationRequired", 511.0),
    ("system.net.httpstatuscode.networkauthenticationrequired", 511.0),
    // `System.Globalization.NumberStyles` — the flag ordinals, so a `Parse`
    // overload can READ the styles it is handed instead of guessing from arity.
    // `AllowHexSpecifier` (512) is the highest flag, which is what lets the
    // parse emitters test `styles >= 512` for "this is hexadecimal".
    ("NumberStyles.None", 0.0),
    ("NumberStyles.AllowLeadingWhite", 1.0),
    ("NumberStyles.AllowTrailingWhite", 2.0),
    ("NumberStyles.AllowLeadingSign", 4.0),
    ("NumberStyles.AllowTrailingSign", 8.0),
    ("NumberStyles.AllowParentheses", 16.0),
    ("NumberStyles.AllowDecimalPoint", 32.0),
    ("NumberStyles.AllowThousands", 64.0),
    ("NumberStyles.AllowExponent", 128.0),
    ("NumberStyles.AllowCurrencySymbol", 256.0),
    ("NumberStyles.AllowHexSpecifier", 512.0),
    ("NumberStyles.Integer", 7.0),
    ("NumberStyles.Number", 111.0),
    ("NumberStyles.Float", 167.0),
    ("NumberStyles.Currency", 383.0),
    ("NumberStyles.Any", 511.0),
    ("NumberStyles.HexNumber", 515.0),
    ("numberstyles.none", 0.0),
    ("numberstyles.allowleadingwhite", 1.0),
    ("numberstyles.allowtrailingwhite", 2.0),
    ("numberstyles.allowleadingsign", 4.0),
    ("numberstyles.allowtrailingsign", 8.0),
    ("numberstyles.allowparentheses", 16.0),
    ("numberstyles.allowdecimalpoint", 32.0),
    ("numberstyles.allowthousands", 64.0),
    ("numberstyles.allowexponent", 128.0),
    ("numberstyles.allowcurrencysymbol", 256.0),
    ("numberstyles.allowhexspecifier", 512.0),
    ("numberstyles.integer", 7.0),
    ("numberstyles.number", 111.0),
    ("numberstyles.float", 167.0),
    ("numberstyles.currency", 383.0),
    ("numberstyles.any", 511.0),
    ("numberstyles.hexnumber", 515.0),
    // ⛔ `Profile::lookup_constant` LOWERCASES the key for a case-insensitive
    // language and looks the cased name up verbatim for a case-sensitive one.
    // That is what the cased/lowercase duplicate pairs further down this table
    // are for — the lowercase row serves VB, the cased row serves C#. Every
    // limit above had only the C# spelling, so `Integer.MaxValue` and
    // `Char.MaxValue` resolved to NOTHING in VB: the first rendered empty and
    // the second trapped in `charCodeAt` under `AscW`.
    //
    // The rows below are the lowercase halves, plus the VB type names (`Integer`
    // is Int32's VB alias, not a different type).
    ("int.maxvalue", 2_147_483_647.0),
    ("int.minvalue", -2_147_483_648.0),
    ("integer.maxvalue", 2_147_483_647.0),
    ("integer.minvalue", -2_147_483_648.0),
    ("short.maxvalue", 32_767.0),
    ("short.minvalue", -32_768.0),
    ("int16.maxvalue", 32_767.0),
    ("int16.minvalue", -32_768.0),
    ("int32.maxvalue", 2_147_483_647.0),
    ("int32.minvalue", -2_147_483_648.0),
    ("byte.maxvalue", 255.0),
    ("byte.minvalue", 0.0),
    ("double.maxvalue", f64::MAX),
    ("double.minvalue", -f64::MAX),
    ("double.nan", f64::NAN),
    ("double.positiveinfinity", f64::INFINITY),
    ("double.negativeinfinity", f64::NEG_INFINITY),
    ("float.maxvalue", 3.4028235e38),
    ("float.minvalue", -3.4028235e38),
    ("single.maxvalue", 3.4028235e38),
    ("single.minvalue", -3.4028235e38),
    ("double.epsilon", 5e-324),
    ("float.epsilon", 1.401298464324817e-45),
    ("single.epsilon", 1.401298464324817e-45),
    ("single.nan", f64::NAN),
    ("single.positiveinfinity", f64::INFINITY),
    ("single.negativeinfinity", f64::NEG_INFINITY),
    // The CODE UNIT, not the one-character string .NET's `Char.MaxValue` really
    // is — this table is f64-only, and every measured use is a limit comparison
    // or an `AscW`.
    ("char.maxvalue", 65535.0),
    ("char.minvalue", 0.0),
    // ⛔ `Long`/`ULong`/`Decimal` limits are deliberately ABSENT. Their maxima
    // are not representable in f64: `Int64.MaxValue` would come back as
    // ...808 rather than ...807. A missing constant is a loud failure; a
    // silently-off-by-one one is not.
    ("commandtype.text", 1.0),
    ("CommandType.Text", 1.0),
    ("commandtype.storedprocedure", 4.0),
    ("CommandType.StoredProcedure", 4.0),
    ("connectionstate.closed", 0.0),
    ("ConnectionState.Closed", 0.0),
    ("connectionstate.open", 1.0),
    ("ConnectionState.Open", 1.0),
    ("regexoptions.none", 0.0),
    ("RegexOptions.None", 0.0),
    ("system.text.regularexpressions.regexoptions.none", 0.0),
    ("System.Text.RegularExpressions.RegexOptions.None", 0.0),
    ("regexoptions.ignorecase", 1.0),
    ("RegexOptions.IgnoreCase", 1.0),
    ("msgboxstyle.okonly", 0.0),
    ("MsgBoxStyle.OkOnly", 0.0),
    ("microsoft.visualbasic.msgboxstyle.okonly", 0.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.OkOnly", 0.0),
    ("msgboxstyle.okcancel", 1.0),
    ("MsgBoxStyle.OkCancel", 1.0),
    ("microsoft.visualbasic.msgboxstyle.okcancel", 1.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.OkCancel", 1.0),
    ("msgboxstyle.abortretryignore", 2.0),
    ("MsgBoxStyle.AbortRetryIgnore", 2.0),
    ("microsoft.visualbasic.msgboxstyle.abortretryignore", 2.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.AbortRetryIgnore", 2.0),
    ("msgboxstyle.yesnocancel", 3.0),
    ("MsgBoxStyle.YesNoCancel", 3.0),
    ("microsoft.visualbasic.msgboxstyle.yesnocancel", 3.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.YesNoCancel", 3.0),
    ("msgboxstyle.yesno", 4.0),
    ("MsgBoxStyle.YesNo", 4.0),
    ("microsoft.visualbasic.msgboxstyle.yesno", 4.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.YesNo", 4.0),
    ("msgboxstyle.retrycancel", 5.0),
    ("MsgBoxStyle.RetryCancel", 5.0),
    ("microsoft.visualbasic.msgboxstyle.retrycancel", 5.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.RetryCancel", 5.0),
    ("msgboxresult.ok", 1.0),
    ("MsgBoxResult.Ok", 1.0),
    ("microsoft.visualbasic.msgboxresult.ok", 1.0),
    ("Microsoft.VisualBasic.MsgBoxResult.Ok", 1.0),
    (
        "system.text.regularexpressions.regexoptions.ignorecase",
        1.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.IgnoreCase",
        1.0,
    ),
    ("regexoptions.multiline", 2.0),
    ("RegexOptions.Multiline", 2.0),
    ("system.text.regularexpressions.regexoptions.multiline", 2.0),
    ("System.Text.RegularExpressions.RegexOptions.Multiline", 2.0),
    ("regexoptions.explicitcapture", 4.0),
    ("RegexOptions.ExplicitCapture", 4.0),
    (
        "system.text.regularexpressions.regexoptions.explicitcapture",
        4.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.ExplicitCapture",
        4.0,
    ),
    ("regexoptions.compiled", 8.0),
    ("RegexOptions.Compiled", 8.0),
    ("system.text.regularexpressions.regexoptions.compiled", 8.0),
    ("System.Text.RegularExpressions.RegexOptions.Compiled", 8.0),
    ("regexoptions.singleline", 16.0),
    ("RegexOptions.Singleline", 16.0),
    (
        "system.text.regularexpressions.regexoptions.singleline",
        16.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.Singleline",
        16.0,
    ),
    ("regexoptions.ignorepatternwhitespace", 32.0),
    ("RegexOptions.IgnorePatternWhitespace", 32.0),
    (
        "system.text.regularexpressions.regexoptions.ignorepatternwhitespace",
        32.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace",
        32.0,
    ),
    ("regexoptions.righttoleft", 64.0),
    ("RegexOptions.RightToLeft", 64.0),
    (
        "system.text.regularexpressions.regexoptions.righttoleft",
        64.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.RightToLeft",
        64.0,
    ),
    ("regexoptions.ecmascript", 256.0),
    ("RegexOptions.ECMAScript", 256.0),
    (
        "system.text.regularexpressions.regexoptions.ecmascript",
        256.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.ECMAScript",
        256.0,
    ),
    ("regexoptions.cultureinvariant", 512.0),
    ("RegexOptions.CultureInvariant", 512.0),
    (
        "system.text.regularexpressions.regexoptions.cultureinvariant",
        512.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.CultureInvariant",
        512.0,
    ),
    // `System.IO.SeekOrigin` — `Begin`/`Current`/`End`, the values
    // `MemoryStream.Seek` branches on.
    ("seekorigin.begin", 0.0),
    ("SeekOrigin.Begin", 0.0),
    ("system.io.seekorigin.begin", 0.0),
    ("System.IO.SeekOrigin.Begin", 0.0),
    ("seekorigin.current", 1.0),
    ("SeekOrigin.Current", 1.0),
    ("system.io.seekorigin.current", 1.0),
    ("System.IO.SeekOrigin.Current", 1.0),
    ("seekorigin.end", 2.0),
    ("SeekOrigin.End", 2.0),
    ("system.io.seekorigin.end", 2.0),
    ("System.IO.SeekOrigin.End", 2.0),
];

static KNOWN_TYPE_MAPPINGS: LazyLock<Vec<KnownTypeMapping>> = LazyLock::new(|| {
    super::component_classes::class_exports()
        .iter()
        .filter_map(|export| {
            let target = export.class.constructor()?.backing.as_ref()?;
            Some(KnownTypeMapping {
                name: leak_string(export.class.name.to_lowercase()),
                interface: export.interface,
                display_name: leak_string(export.class.name.clone()),
                target: match target {
                    ConstructorTarget::Host(target) => KnownTypeTarget::Host {
                        module: leak_string(target.module.clone()),
                        constructor: leak_string(target.name.clone()),
                    },
                    ConstructorTarget::Common(name) => KnownTypeTarget::Common {
                        emit: leak_string(name.clone()),
                    },
                },
            })
        })
        .collect()
});

fn leak_string(value: String) -> &'static str {
    Box::leak(value.into_boxed_str())
}

pub fn known_type_mappings() -> &'static [KnownTypeMapping] {
    KNOWN_TYPE_MAPPINGS.as_slice()
}

pub fn is_known_constant(name: &str) -> bool {
    known_constants().contains(&name)
}

pub fn known_constants() -> &'static [&'static str] {
    KNOWN_CONSTANTS
}

pub fn namespace_constants() -> &'static [(&'static str, f64)] {
    NAMESPACE_CONSTANTS
}

pub fn capitalize_data_type(name: &str) -> String {
    match name {
        "dataset" => "DataSet",
        "datatable" => "DataTable",
        "dataadapter" => "DataAdapter",
        _ => return String::new(),
    }
    .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_known_type_mappings_exclude_winforms_entries() {
        assert!(
            known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "stringbuilder")
        );
        assert!(
            !known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "form")
        );
    }
}
