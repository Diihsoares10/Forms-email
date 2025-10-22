from flask import Flask, request, render_template
import pythoncom
import win32com.client as win32

app = Flask(__name__)

REMETENTE = "md.diego@hotmail.com"  # sua conta do Outlook

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/send_email", methods=["POST"])
def send_email():
    pythoncom.CoInitialize()

    nome = request.form.get("nome")
    sobrenome = request.form.get("sobrenome")
    idade = request.form.get("idade") or "Não informado"
    profissao = request.form.get("profissao") or "Não informado"
    genero = request.form.get("genero") or "Não informado"
    email = request.form.get("email")  # destinatário dinâmico
    telefone = request.form.get("telefone") or "Não informado"
    mensagem_usuario = request.form.get("mensagem")
    newsletter = "Sim" if request.form.get("newsletter") else "Não"

    corpo_email = f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #f1c40f;">Nova mensagem de contato</h2>
        <ul style="list-style-type: none; padding: 0;">
            <li><strong>Nome:</strong> {nome} {sobrenome}</li>
            <li><strong>Idade:</strong> {idade}</li>
            <li><strong>Profissão:</strong> {profissao}</li>
            <li><strong>Gênero:</strong> {genero}</li>
            <li><strong>E-mail:</strong> {email}</li>
            <li><strong>Telefone:</strong> {telefone}</li>
            <li><strong>Deseja receber newsletter:</strong> {newsletter}</li>
        </ul>
        <p><strong>Mensagem:</strong><br>{mensagem_usuario}</p>
    </div>
    """

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        accounts = outlook.Session.Accounts
        for account in accounts:
            if account.SmtpAddress.lower() == REMETENTE.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                break

        mail.To = email  # envia para o e-mail digitado
        mail.Subject = f"Contato do site: {nome} {sobrenome}"
        mail.HTMLBody = corpo_email
        mail.Send()

        return render_template("resultado.html", mensagem=f"✅ E-mail enviado com sucesso para {email}!")
    except Exception as e:
        return render_template("resultado.html", mensagem=f"❌ Erro ao enviar e-mail: {e}")

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
