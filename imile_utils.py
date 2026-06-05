import re

def login_imile(page, usuario, senha):
    """
    Realiza o login no portal da iMile.
    """
    print("Acessando o portal iMile...")
    page.goto("https://ds-login.imile.com/", wait_until="domcontentloaded")
    
    page.fill('input[type="text"]', str(usuario or ""))
    page.fill('input[type="password"]', str(senha or ""))
    page.get_by_role("button", name=re.compile(r"登录|Conecte-se|Login", re.IGNORECASE)).click()
    
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(5000)

    try:
        page.locator('.close-icon').first.click(force=True)
    except:
        pass
    page.wait_for_timeout(3000)
    print("Login na iMile realizado com sucesso.")
