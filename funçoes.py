import main
import smtplib
import email.message


def enviar_mensagem():
    print(main.loja)
    msg = email.message.Message()
    msg["Subject"] = f'Onepage do dia {main.dia_indicador.day}/{main.dia_indicador.month}-{main.loja}'
    msg["From"] = "delsangue666@gmail.com"
    msg["To"] = "joaocarlosreis99@gmail.com"
    # msg["To"] = main.email.loc[main.email['Loja']== main.loja, 'E-mail'].value[0]
    msg["Cc"] = "joaocarlosreis99+copia@gmail.com"
    msg["Bcc"] = "delsangue666@gmail.com"
    corpo_email = f"""<p>Olá, Del Scória aqui</p>
        <h2>Bom dia {main.loja}</h2>
        <p>O Resultado de ontem {main.dia_indicador.day}/{main.dia_indicador.month} da Loja <strong>{main.loja}</strong> foi...</p>
        <p>Faturamento do dia da loja {main.loja} R$:{main.faturamento_dia:,.2f}</p>
        <p>Quantidade de produtos vendido por dia Unidade = [{main.quant_produtos_dia}]</p>
        <p>Tikte médio por dia R$:{main.tikte_medio_dia:,.2f}</p>
        <p>Esse são os dados do dia anterior, tenha um òtimo dia e boas vendas</p>
        """
    corpo_email = corpo_email.encode("utf-8")
    msg.add_header("Content-Type", "text/html")
    msg.set_payload(corpo_email)
    servidor = smtplib.SMTP('smtp.gmail.com:587')
    servidor.starttls()
    servidor.login(msg["From"], "ioor tbkh zhrt yuei")
    servidor.send_message(msg)
    servidor.quit()
    print(f'mensagem da loja {main.loja} foi enviada com sucesso!')
    print('#'* 50)



