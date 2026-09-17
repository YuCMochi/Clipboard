fn main() {
    // 版本資訊從 Cargo.toml 傳給 app.rc 的 VERSIONINFO，版本號只維護一處
    let num = |key| std::env::var(key).unwrap();
    embed_resource::compile(
        "app.rc",
        [
            format!("VER_MAJOR={}", num("CARGO_PKG_VERSION_MAJOR")),
            format!("VER_MINOR={}", num("CARGO_PKG_VERSION_MINOR")),
            format!("VER_PATCH={}", num("CARGO_PKG_VERSION_PATCH")),
            format!("VER_STR=\"{}\"", num("CARGO_PKG_VERSION")),
            format!("VER_AUTHOR=\"{}\"", num("CARGO_PKG_AUTHORS")),
        ],
    );
}
