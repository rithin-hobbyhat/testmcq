import os
import shutil
import stat
import subprocess
import sys
import tempfile

def remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass

def main():
    msg = sys.argv[1] if len(sys.argv) > 1 else 'Deploy MCQ update to hobbyhat'
    src = os.path.abspath('web')
    temp_dir = tempfile.mkdtemp(prefix='hobbyhat_deploy_')

    try:
        print('Cloning hobbyhat into temporary workspace...')
        subprocess.run(['git', 'clone', '--depth', '1', 'https://github.com/rithin-hobbyhat/hobbyhat.git', temp_dir], check=True)

        mcq_dest = os.path.join(temp_dir, 'mcq')
        os.makedirs(mcq_dest, exist_ok=True)

        for item in os.listdir(src):
            s = os.path.join(src, item)
            d = os.path.join(mcq_dest, item)
            if os.path.isdir(s):
                if os.path.exists(d):
                    shutil.rmtree(d, onerror=remove_readonly)
                shutil.copytree(s, d)
            else:
                shutil.copy2(s, d)

        subprocess.run(['git', '-C', temp_dir, 'add', '.'], check=True)
        status = subprocess.run(['git', '-C', temp_dir, 'status', '--porcelain'], capture_output=True, text=True).stdout
        if status.strip():
            subprocess.run(['git', '-C', temp_dir, 'commit', '-m', msg], check=True)
            subprocess.run(['git', '-C', temp_dir, 'push', 'origin', 'main'], check=True)
            print('Pushed update to hobbyhat successfully!')
        else:
            print('hobbyhat/mcq is already up to date.')
    finally:
        shutil.rmtree(temp_dir, onerror=remove_readonly, ignore_errors=True)

if __name__ == '__main__':
    main()
