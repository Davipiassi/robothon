from botcity.maestro import *
from email_sender import EmailSender
from jinja2 import Environment, FileSystemLoader
import re
from datetime import datetime
import pandas as pd

BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    try:
        maestro = BotMaestroSDK.from_sys_args()
        execution = maestro.get_execution()
        task_info = maestro.get_task(maestro.task_id)
        input_path = task_info.parameters.get('report path')
        notification_receiver = task_info.parameters.get('notification email')

        maestro.alert(
            task_id=execution.task_id,
            title='Bot Iniciou',
            message='O robô acordou para te ajudar!',
            alert_type=AlertType.INFO
        )
    
        report_content = read_report_file(input_path)

        pan_codes = re.findall(r'PAN\s+([0-9]{19})', report_content)
        adv_reasons = re.findall(r'ADV REASON\s+(.{8})\s+SWCH DATE', report_content)
        org_dates = re.findall(r'ORG DATE\s+([0-9]{2}-[0-9]{2})', report_content)
        org_amounts = re.findall(r'ORG AMT\s+([0-9]{1,2}\.[0-9]{2})', report_content)
        acq_inst_names = re.findall(r'ACQ INST NAME\s+(.{30})', report_content)

        for i in range(len(acq_inst_names)):
            acq_inst_names[i] = acq_inst_names[i].rstrip()

        report_data = {
            'PAN': pan_codes,
            'ADV_REASON': adv_reasons,
            'ORG_DATE': org_dates,
            'ORG_AMT': org_amounts,
            'ACQ_INST_NAME': acq_inst_names
        }
        
        maestro.alert(
            task_id=execution.task_id,
            title='Itens lidos',
            message='O robô conseguiu ler o conteúdo do arquivo e está preparando uma resposta para você!',
            alert_type=AlertType.INFO
        )

        output_path = r'resources\layout-reapresentacao.xlsx'
        generate_excel(output_path, report_data)
        post_excel_artifact(maestro, execution.task_id, output_path)
        
        email = EmailSender('asimovreportbot@gmail.com', 'jltw nmpz suos romk')
        email.connect()
        email_content = setup_html_template('email.html', {
            'current_date': datetime.now(),
            'responses_count': len(report_data['PAN'])
        })            
        email.send_email(
            receivers=[notification_receiver],
            subject='Relatório MasterCard processado!',
            content=email_content,
            files=[output_path]
        )
        email.disconnect()
        
        maestro.alert(
            task_id=execution.task_id,
            title='Stakeholders notificados',
            message=f'{notification_receiver} foi avisado de que um novo Layout de Reapresentação está disponível.',
            alert_type=AlertType.INFO
        )
        
        maestro.finish_task(
            task_id=execution.task_id,
            status=AutomationTaskFinishStatus.SUCCESS,
            message='Tarefa finalizada com sucesso.',
            total_items=len(pan_codes),
            processed_items=len(report_data['PAN']),
            failed_items=len(pan_codes) - len(report_data['PAN'])
        )
    except FileNotFoundError as e:
        maestro.alert(
            task_id=execution.task_id,
            title='O arquivo não foi encontrado!',
            message=e,
            alert_type=AlertType.ERROR
        )
    
    except Exception as e:
        maestro.alert(
            task_id=execution.task_id,
            title='Execução interrompida por erro!',
            message=e,
            alert_type=AlertType.ERROR
        )

def not_found(label):
    print(f'Element not found: {label}')

def read_report_file(path):
    with open(path, 'r') as file:
        return file.read()
        
def generate_excel(path, data):
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)

def post_excel_artifact(maestro, task_id, path):
    maestro.post_artifact(
        task_id=task_id,
        artifact_name='Layout de Reapresentação.xlsx',
        filepath=path
    )            
    
def setup_html_template(template, arguments):
    env = Environment(loader=FileSystemLoader('resources/templates'))
    template = env.get_template(template)
    return template.render(arguments)

if __name__ == '__main__':
    main()