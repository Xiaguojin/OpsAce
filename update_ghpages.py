import subprocess, os, tempfile, shutil

os.chdir('C:/Users/guojin.xia/WorkBuddy/2026-07-27-14-04-47')

# Read latest dashboard
with open('dashboard_with_permissions.html', 'r', encoding='utf-8') as f:
    content = f.read()
with open('echarts.min.js', 'rb') as f:
    echarts_content = f.read()

# Set up gh-pages index
subprocess.run(['git', 'read-tree', 'gh-pages'], check=True)

# Write blobs
tmpdir = tempfile.mkdtemp()
with open(os.path.join(tmpdir, 'index.html'), 'w', encoding='utf-8') as f:
    f.write(content)
with open(os.path.join(tmpdir, 'echarts.min.js'), 'wb') as f:
    f.write(echarts_content)

def write_blob(path):
    with open(path, 'rb') as f:
        result = subprocess.run(
            ['git', 'hash-object', '-w', '--stdin'],
            input=f.read(), capture_output=True
        )
    return result.stdout.decode().strip()

index_blob = write_blob(os.path.join(tmpdir, 'index.html'))
echarts_blob = write_blob(os.path.join(tmpdir, 'echarts.min.js'))
print(f'index: {index_blob}')
print(f'echarts: {echarts_blob}')

subprocess.run(
    ['git', 'update-index', '--add', '--cacheinfo', '100644', index_blob, 'index.html'],
    check=True
)
subprocess.run(
    ['git', 'update-index', '--add', '--cacheinfo', '100644', echarts_blob, 'echarts.min.js'],
    check=True
)

tree = subprocess.check_output(['git', 'write-tree']).decode().strip()
gh_head = subprocess.check_output(['git', 'rev-parse', 'gh-pages']).decode().strip()

commit = subprocess.run(
    ['git', 'commit-tree', tree, '-p', gh_head],
    input=b'Update dashboard to 169f187 (login fallback + fix)',
    capture_output=True
).stdout.decode().strip()

subprocess.run(
    ['git', 'update-ref', 'refs/heads/gh-pages', commit, gh_head],
    check=True
)

# Restore main
subprocess.run(['git', 'read-tree', 'main'], check=True)
shutil.rmtree(tmpdir)
print(f'gh-pages updated: {commit}')
