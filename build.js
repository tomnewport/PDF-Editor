// Build script: bundles the renderer with esbuild and copies static assets
// (HTML, CSS) plus the pdf.js worker into the dist/ directory that Electron loads.
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const root = __dirname;
const dist = path.join(root, 'dist');

function copy(from, to) {
  fs.copyFileSync(from, to);
  console.log(`copied ${path.relative(root, from)} -> ${path.relative(root, to)}`);
}

function copyDir(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    const src = path.join(from, entry.name);
    const dest = path.join(to, entry.name);
    if (entry.isDirectory()) copyDir(src, dest);
    else copy(src, dest);
  }
}

async function main() {
  fs.mkdirSync(dist, { recursive: true });

  await esbuild.build({
    entryPoints: [path.join(root, 'src', 'renderer.js')],
    bundle: true,
    outfile: path.join(dist, 'renderer.js'),
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    sourcemap: true,
    logLevel: 'info',
  });

  // Static assets
  copy(path.join(root, 'public', 'index.html'), path.join(dist, 'index.html'));
  copy(path.join(root, 'public', 'styles.css'), path.join(dist, 'styles.css'));
  copy(path.join(root, 'assets', 'icon.svg'), path.join(dist, 'icon.svg'));
  copy(path.join(root, 'assets', 'icon.png'), path.join(dist, 'icon.png'));

  // pdf.js worker (loaded as a module worker at runtime)
  copy(
    path.join(root, 'node_modules', 'pdfjs-dist', 'build', 'pdf.worker.min.mjs'),
    path.join(dist, 'pdf.worker.min.mjs')
  );
  copyDir(
    path.join(root, 'node_modules', 'pdfjs-dist', 'standard_fonts'),
    path.join(dist, 'standard_fonts')
  );
  copyDir(
    path.join(root, 'node_modules', 'pdfjs-dist', 'web', 'images'),
    path.join(dist, 'images')
  );

  console.log('build complete');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
