"""
Developer:Luis'Guillermo
Date:11/2022
Version: 1.0
Python: 3.10

"""
import csv
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from os import system
from os.path import exists
import conn

hora = datetime.now()
log_str = hora.strftime('%Y%m%d_%H%M')
archivo_logsend = f'sendmail_log{log_str}.log'
archivo_logerror = f'error_log{log_str}.log'

if not exists('logs'):
    system('mkdir logs')

if not exists('datos.csv'):
    with open(f'logs/{archivo_logerror}', 'w') as log:
        log.write(f'{datetime.now()} ERROR: No existe el archivo datos.csv en la ruta, ¡Verifique porfavor! \n')
        print('No existe el archivo datos.csv en la ruta, ¡Verifique porfavor!')
        exit(0)
with open(f'logs/{archivo_logsend}', 'w') as log:
    log.write(f'{datetime.now()}: Inicio ejecucion del programa\n')

    #Funciones
    def enviar():
        #Parametros mailserver
        email_emisor = conn.email_emisor
        email_contrasena = conn.email_contrasena
        asunto = conn.asunto
        mailserver = conn.mailserver
        msport = conn.msport

        em = EmailMessage()
        #Aqui rescatamos las variables que contienen el cuerpo del mensaje con los datos del cliente
        em['From'] = email_emisor
        em['To'] = correo
        em['Subject'] = asunto
        em.set_content(html, subtype="html")

        contexto = ssl.create_default_context()

        #Loguea y envia el mensaje
        with smtplib.SMTP_SSL(mailserver, msport, context=contexto) as smtp:
            #Si necesita depurar algun error al momento del envio que no aparezca en el archivo error_log, descomente la siguiente opcion y revise la consola.
            #smtp.set_debuglevel(debuglevel=2)

            try:
                login = smtp.login(email_emisor, email_contrasena)  # se authentica para enviar.
                log.write(f'{datetime.now()}: {login} \n')
            except smtplib.SMTPAuthenticationError as serr:
                log.write(f'{datetime.now()} ERROR: {serr} \n')
                login = None
            if login is not None:

                smtp.sendmail(email_emisor, correo, em.as_string()) #Se envia el correo.

                log.write(f'{datetime.now()}: Mail enviado a {correo} \n')



    with open('datos.csv', 'r') as archivo:
        lector = csv.reader(archivo, delimiter=",")
        # omitir el encabezado
        next(lector, None)

        for fila in lector:
            # Aqui leemos la lista linea por linea
            nombre = fila[0]
            correo = fila[1]
            usuario = fila[2]
            contraseña = fila[3]

            html = (f"""\
                                        <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">
        
                            <head>
                                <title></title>
                                <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                <!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]-->
                                <!--[if !mso]><!-->
                                <link href="https://fonts.googleapis.com/css?family=Abril+Fatface" rel="stylesheet" type="text/css">
                                <link href="https://fonts.googleapis.com/css?family=Alegreya" rel="stylesheet" type="text/css">
                                <link href="https://fonts.googleapis.com/css?family=Arvo" rel="stylesheet" type="text/css">
                                <link href="https://fonts.googleapis.com/css?family=Bitter" rel="stylesheet" type="text/css">
                                <link href="https://fonts.googleapis.com/css?family=Cabin" rel="stylesheet" type="text/css">
                                <link href="https://fonts.googleapis.com/css?family=Ubuntu" rel="stylesheet" type="text/css">
                                <!--<![endif]-->
        
                            </head>
        
                            <body style="background-color: #FFFFFF; margin: 0; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
                                <table class="nl-container" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #FFFFFF;">
                                    <tbody>
                                        <tr>
                                            <td>
                                                <table class="row row-1" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f5f5f5;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 0px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                                <table class="image_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad" style="padding-bottom:10px;width:100%;padding-right:0px;padding-left:0px;">
                                                                                            <div class="alignment" align="center" style="line-height:10px"><img src="http://www.moscoso.cl/wp-content/uploads/2022/08/cropped-logo_adm-e1493820237612.jpg" style="display: block; height: auto; border: 0; width: 125px; max-width: 100%;" width="125" alt="your-logo" title="your-logo"></div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <table class="row row-2" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f5f5f5;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #ffffff; color: #000000; width: 500px;" width="500">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 15px; padding-bottom: 20px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                                <table class="image_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad" style="padding-bottom:5px;padding-left:5px;padding-right:5px;width:100%;">
                                                                                            <div class="alignment" align="center" style="line-height:10px"><img class="big" src="http://www.moscoso.cl/wp-content/uploads/2022/10/gif-resetpass.gif" style="display: block; height: auto; border: 0; width: 350px; max-width: 100%;" width="350" alt="reset-password" title="reset-password"></div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="heading_block block-2" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad" style="text-align:center;width:100%;">
                                                                                            <h1 style="margin: 0; color: #393d47; direction: ltr; font-family: Tahoma, Verdana, Segoe, sans-serif; font-size: 25px; font-weight: 400; letter-spacing: normal; line-height: 120%; text-align: center; margin-top: 0; margin-bottom: 0;"><strong>Estimad@ {nombre}</strong></h1>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="text_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family: Tahoma, Verdana, sans-serif">
                                                                                                <div class style="font-size: 12px; font-family: Tahoma, Verdana, Segoe, sans-serif; mso-line-height-alt: 18px; color: #393d47; line-height: 1.5;">
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 18px;">Junto con saludar, manteniendo los protocolos de seguridad, es importante ir cambiando las credenciales de acceso cada cierto tiempo minimizando los riesgos de accesos no autorizados a los sistemas,</p>
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 18px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dado lo anterior queremos informar que el lunes 07 de Noviembre 2022 se realizara el cambio de contraseña de acceso a la VPN (FortiClient) de su usuario, el motivo de este correo es para ponerle en conocimiento.</p>
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 18px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Los nuevos accesos comenzaran a regir desde la fecha antes mencionada, se avisa con anticipación con el ánimo no interrumpir sus labores.</p>
                                                                                                </div>
                                                                                            </div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="text_block block-4" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                                    <tr>
                                                                                        <td class="pad" style="padding-bottom:5px;padding-left:10px;padding-right:10px;padding-top:10px;">
                                                                                            <div style="font-family: Tahoma, Verdana, sans-serif">
                                                                                                <div class style="font-size: 12px; font-family: Tahoma, Verdana, Segoe, sans-serif; text-align: center; mso-line-height-alt: 18px; color: #393d47; line-height: 1.5;">
                                                                                                    <p style="margin: 0; mso-line-height-alt: 18px;">Credenciales</p>
                                                                                                    <p style="margin: 0; mso-line-height-alt: 18px;">&nbsp;Usuario: {usuario}</p>
                                                                                                    <p style="margin: 0; mso-line-height-alt: 18px;">&nbsp; &nbsp; &nbsp;Contraseña: {contraseña}</p>
                                                                                                </div>
                                                                                            </div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <table class="row row-3" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f5f5f5;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                                <table class="text_block block-1" width="100%" border="0" cellpadding="15" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family: Tahoma, Verdana, sans-serif">
                                                                                                <div class style="font-size: 12px; font-family: Tahoma, Verdana, Segoe, sans-serif; mso-line-height-alt: 14.399999999999999px; color: #393d47; line-height: 1.2;">
                                                                                                    <p style="margin: 0; font-size: 14px; text-align: center; mso-line-height-alt: 16.8px;"><br><span style="font-size:10px;">por favor siéntase libre de contactarnos en operaciones@moscoso.cl </span></p>
                                                                                                </div>
                                                                                            </div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <table class="row row-4" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #fff;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                                <table class="html_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family:Arial, Helvetica Neue, Helvetica, sans-serif;text-align:center;" align="center"><div style="height:30px;">&nbsp;</div></div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="html_block block-2" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family:Arial, Helvetica Neue, Helvetica, sans-serif;text-align:center;" align="center"><div style="margin-top: 25px;border-top:1px dashed #D6D6D6;margin-bottom: 20px;"></div></div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="text_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family: Tahoma, Verdana, sans-serif">
                                                                                                <div class style="font-size: 12px; font-family: Tahoma, Verdana, Segoe, sans-serif; mso-line-height-alt: 14.399999999999999px; color: #C0C0C0; line-height: 1.2;">
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 14.399999999999999px;">&nbsp;</p>
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 14.399999999999999px;">&nbsp;</p>
                                                                                                    <p style="margin: 0; text-align: center; mso-line-height-alt: 14.399999999999999px;">Cruz del sur 133 Of 802, Las Condes&nbsp; /&nbsp;&nbsp;Contacto@moscoso.cl / (+56) 22 570 5900<a href="http://www.example.com" style></a></p>
                                                                                                    <p style="margin: 0; font-size: 12px; text-align: center; mso-line-height-alt: 14.399999999999999px;"><span style="color:#c0c0c0;">&nbsp;</span></p>
                                                                                                </div>
                                                                                            </div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <table class="html_block block-4" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad">
                                                                                            <div style="font-family:Arial, Helvetica Neue, Helvetica, sans-serif;text-align:center;" align="center"><div style="height-top: 20px;">&nbsp;</div></div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <table class="row row-5" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                                <table class="icons_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="pad" style="vertical-align: middle; color: #9d9d9d; font-family: inherit; font-size: 15px; padding-bottom: 5px; padding-top: 5px; text-align: center;">
                                                                                            <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                                <tr>
                                                                                                    <td class="alignment" style="vertical-align: middle; text-align: center;">
                                                                                                        <!--[if vml]><table align="left" cellpadding="0" cellspacing="0" role="presentation" style="display:inline-block;padding-left:0px;padding-right:0px;mso-table-lspace: 0pt;mso-table-rspace: 0pt;"><![endif]-->
                                                                                                        <!--[if !vml]><!-->
                                                                                                        <table class="icons-inner" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block; margin-right: -4px; padding-left: 0px; padding-right: 0px;" cellpadding="0" cellspacing="0" role="presentation">
                                                                                                            <!--<![endif]-->
        
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table><!-- End -->
                            </body>
        
                            </html>
                                    """)
            enviar()
            log.write(f'{datetime.now()}: Proceso de envio finalizado. \n')
            print('Termino el proceso.')



