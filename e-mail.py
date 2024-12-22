import win32com.client as client
import pathlib

outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = 'Jonas.Komischke@edu.bib.de'

message.Display()

message.Subject = 'Travel Guide'

hauptbild_path = pathlib.Path("Group 16.png")
hauptbild_absolute = str(hauptbild_path.absolute())


image = message.Attachments.Add(hauptbild_absolute)

SriLanka_path = pathlib.Path("SriLanka.png")
SriLanka_absolute = str(SriLanka_path.absolute())


image_2 = message.Attachments.Add(SriLanka_absolute)
 

html_body = """
<div style="text-align: center; background-color: #fff;">
    <img src="cid:hauptbild-img" style="max-width: 100%; height: auto;">
    <table style="margin: 0 auto; text-align: center; width: 600px; padding: 20px 40px 52px;" border="0" cellspacing="0">
        <tr>
            <td colspan="2" style="padding:0; padding-top: 20px; padding-left: 30px; padding-right: 30px;">
                <h1 style="color: #FF3465; font-size: 30px;">Hallo Anika,</h1>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="padding-top: 10px;">
                <p style="font-size: 15px; max-width: 600px;">
                Stell dir vor, du könntest eine Auszeit vom Alltag nehmen, <span style="font-weight:bold; color: #FF3465; ">neue</span> Orte entdecken und einfach mal die Seele baumeln lassen. Ein Urlaub bietet dir die <span style="font-weight:bold; color: #FF3465;">perfekte Gelegenheit</span>, den Stress hinter dir zu lassen, neue Energie zu tanken und unvergessliche Erinnerungen zu sammeln. Es gibt so viele faszinierende Ziele, von idyllischen Stränden bis zu aufregenden Städten, die nur darauf warten, von dir entdeckt zu werden. Warum also nicht den Schritt wagen und dir die verdiente Erholung gönnen? Dein Körper und Geist werden es dir danken!
                </p>
            </td>
        </tr>
        <tr style="background-color: #1A2D72; text-align: justify; color:#fff">
            <td style="width: 50%; padding-left: 30px; padding-right: 30px;">
                <img src="cid:SriLanka-img" style="max-width: 100%; height: auto;">
            </td>
            <td style="width: 50%; padding-left: 30px; padding-right: 30px;">
                <h2 style="text-decoration: underline;">Sri-Lanka</h2>
                <p>Stell dir vor, du wachst auf, umgeben von sanften Wellen und dem beruhigenden Rauschen der Palmen im Wind. Die warme, tropische Brise streichelt deine Haut, während die ersten Sonnenstrahlen den Himmel in sanften Pastellfarben erleuchten. Du befindest dich an einem paradiesischen</p>
            </td>
        </tr>
        <tr style="background-color: #1A2D72;  text-align: justify; color: #fff;">
            <td colspan="2" style="padding-top: 11px; padding-left: 30px; padding-right: 30px;">
                <p>
                Der Duft von salziger Meeresluft und exotischen Blumen erfüllt die Luft, während du dich dem entspannten Rhythmus der Insel hingibst. Morgens kannst du barfuß am Strand entlang spazieren, den Ozean in seiner ganzen Pracht genießen und dabei den Sonnenaufgang bestaunen. Tagsüber locken erfrischende Badeerlebnisse im klaren Wasser und die Möglichkeit, an der Küste zu schnorcheln oder zu tauchen, um die farbenprächtige Unterwasserwelt zu entdecken.
                </p>
            </td>
        </tr>
        <tr style="background-color: #1A2D72;">
            <td colspan="2" style="padding-left: 30px; padding-right: 30px; padding-bottom: 40px;">
                 <a href="https://www.lingling.bplaced.net" style="text-decoration: none; color: #fff;">
                     <button type="button" style="width: 148px; height: 38px; background-color: #FF3465; border-radius: 10px;">Let's go!</button>
                 </a>
            </td>
        </tr>
    </table>
</div>

    """

image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "hauptbild-img")
image_2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "SriLanka-img")
message.HTMLBody = html_body
