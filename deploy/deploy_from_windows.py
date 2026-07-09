# -*- coding: utf-8 -*-
"""
本机一键部署:上传代码包+数据包到腾讯云并安装启动。

用法(服务器地址与密码从环境变量读,不进命令行历史):
    set QUOTE_SSH_HOST=你的服务器IP
    set QUOTE_SSH_PASS=你的SSH密码
    python deploy/deploy_from_windows.py
"""
import os, sys, time
import paramiko

HOST = os.environ.get("QUOTE_SSH_HOST", "YOUR_SERVER_IP")
USER = os.environ.get("QUOTE_SSH_USER", "ubuntu")
DIST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist")


def run(ssh, cmd, timeout=600):
    print(f"\n$ {cmd}")
    stdin, stdout, stderr = ssh.exec_command(cmd, timeout=timeout, get_pty=True)
    out = stdout.read().decode("utf-8", "replace")
    code = stdout.channel.recv_exit_status()
    print(out[-4000:] if len(out) > 4000 else out)
    if code != 0:
        print(f"!! 退出码 {code}")
    return code, out


def main():
    password = os.environ.get("QUOTE_SSH_PASS")
    if not password:
        print("请先设置环境变量 QUOTE_SSH_PASS")
        sys.exit(1)
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    print(f"连接 {USER}@{HOST} ...")
    ssh.connect(HOST, username=USER, password=password, timeout=20)

    sftp = ssh.open_sftp()
    for name in ("quote-app.tar.gz", "quote-data.tar.gz"):
        src = os.path.join(DIST, name)
        size = os.path.getsize(src)
        print(f"上传 {name} ({size/1e6:.1f} MB) ...")
        t0 = time.time()
        sftp.put(src, f"/home/ubuntu/{name}")
        print(f"  完成 {time.time()-t0:.0f}s")
    sftp.close()

    # docker 权限探测:不在 docker 组则用 sudo -n
    code, _ = run(ssh, "docker info >/dev/null 2>&1 && echo OK || echo NEED_SUDO")
    prefix = ""
    _, out = run(ssh, "docker info >/dev/null 2>&1 && echo OK || (sudo -n docker info >/dev/null 2>&1 && echo SUDO_OK || echo FAIL)")
    if "SUDO_OK" in out:
        prefix = "sudo "
    elif "OK" not in out:
        print("!! ubuntu 用户无 docker 权限且 sudo 需密码,请先在服务器执行: sudo usermod -aG docker ubuntu")
        sys.exit(2)

    run(ssh, "mkdir -p /home/ubuntu/quote-app && tar -xzf /home/ubuntu/quote-app.tar.gz -C /home/ubuntu/quote-app")
    # deploy_remote.sh 内的 docker 命令按需加 sudo
    if prefix:
        run(ssh, "sed -i 's/^docker /sudo docker /; s/ docker compose/ sudo docker compose/; s/^  docker run/  sudo docker run/' /home/ubuntu/quote-app/deploy_remote.sh")
    code, out = run(ssh, "bash /home/ubuntu/quote-app/deploy_remote.sh", timeout=1200)
    if code == 0 and "DONE" in out:
        print("\n========== 部署成功 ==========")
        print(f"访问地址: https://{HOST}:9443  (自签证书,首次访问点『继续前往』)")
        print("提醒: 腾讯云安全组需放行 TCP 9443")
    else:
        print("\n部署未完成,请把上面输出发给我排查。")
    ssh.close()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
