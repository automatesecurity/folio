#!/usr/bin/env bash
# Cross-compiles a single-file Folio.exe from Linux.
# Needs: mingw-w64, a Rust toolchain, and either the x86_64-pc-windows-gnu target
# or Rust's library source (then std is built from source with -Zbuild-std).
set -euo pipefail
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
TARGET=x86_64-pc-windows-gnu
SHIM="$ROOT/src-tauri/target/wv2-shim"
OUT="$ROOT/dist-portable"
VERSION=$(node -p "require('$ROOT/package.json').version")

cd "$ROOT"
npm run build
(cd src-tauri && cargo fetch --target "$TARGET")
LOADER=$(ls "$HOME"/.cargo/registry/src/*/webview2-com-sys-*/x64/WebView2Loader.dll | tail -1)

mkdir -p "$SHIM" "$OUT"
x86_64-w64-mingw32-gcc -O2 -c "$ROOT/packaging/windows-portable/webview2_shim.c" \
  -DLOADER_PATH="\"$LOADER\"" -o "$SHIM/webview2_shim.o"
rm -f "$SHIM/libWebView2Loader.dll.a"
x86_64-w64-mingw32-ar rcs "$SHIM/libWebView2Loader.dll.a" "$SHIM/webview2_shim.o"

STD=()
if ! rustup target list --installed 2>/dev/null | grep -q "$TARGET"; then
  export RUSTC_BOOTSTRAP=1
  STD=(-Zbuild-std=std,panic_abort)
fi
export RUSTFLAGS="-L native=$SHIM -C target-feature=+crt-static -C link-arg=-static-libgcc"
export CARGO_TARGET_X86_64_PC_WINDOWS_GNU_LINKER=x86_64-w64-mingw32-gcc
(cd src-tauri && cargo build --release --target "$TARGET" --features tauri/custom-protocol ${STD[@]+"${STD[@]}"})
cp "src-tauri/target/$TARGET/release/folio.exe" "$OUT/Folio-$VERSION-portable-x64.exe"
echo "$OUT/Folio-$VERSION-portable-x64.exe"
