from botcity.core import DesktopBot
from botcity.maestro import *
import re
import pandas as pd

BotMaestroSDK.RAISE_NOT_CONNECTED = False

success_count = 0

def main():
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    bot = DesktopBot()

    try:
        report_path = rf"{execution.parameters['report_path']}"

        maestro.alert(
            task_id=execution.task_id,
            title="Bot Iniciou",
            message=f"A automação começou a executar. Arquivo no caminho {report_path}",
            alert_type=AlertType.INFO
        )

        report_content = read_report_file(report_path)
        report_data = extract_data(report_content)

        maestro.alert(
            task_id=execution.task_id,
            title="Registros encontrados",
            message=f"O processamento encontrou {len(report_data['PAN'])} itens",
            alert_type=AlertType.INFO
        )

        excel_path = r'resources\output.xlsx'
        generate_excel(excel_path, report_data)
        post_excel_artifact(maestro, execution.task_id, excel_path)

        maestro.finish_task(
            task_id=execution.task_id,
            status=AutomationTaskFinishStatus.SUCCESS,
            message="Tarefa concluída com sucesso.",
            total_items=len(report_content['PAN']),
            processed_items=success_count
        )
    except Exception as error:
        maestro.alert(
            task_id=execution.task_id,
            title="Falha ao executar",
            message=error,
            alert_type=AlertType.ERROR
        )









def not_found(label):
    print(f"Element not found: {label}")

def read_report_file(path):
    with open(path, 'r') as file:
        return file.read()
    
def extract_data(text_content):
    pan_codes = re.findall(r'PAN\s+([0-9]{19})', text_content)
    adv_reasons = re.findall(r'ADV REASON\s+(.{8})\s+SWCH DATE', text_content)
    org_dates = re.findall(r'ORG DATE\s+([0-9]{2}-[0-9]{2})', text_content)
    org_amounts = re.findall(r'ORG AMT\s+([0-9]{1,2}\.[0-9]{2})', text_content)
    acq_inst_names = re.findall(r'ACQ INST NAME\s+(.{30})', text_content)
    for i in range(len(acq_inst_names)):
        acq_inst_names[i] = acq_inst_names[i].rstrip()
        success_count += 1

    return {
        'PAN': pan_codes,
        'ADV_REASON': adv_reasons,
        'ORG_DATE': org_dates,
        'ORG_AMT': org_amounts,
        'ACQ_INST_NAME': acq_inst_names
    }
    
def generate_excel(path, data):
    df = pd.DataFrame(data)
    header = ['Código PAN', 'Razão da Exceção', 'Data da Compra', 'Valor da Compra', 'Instituição da Compra']
    df.to_excel(path, index=False, header=header)

def post_excel_artifact(maestro, task_id, path):
    maestro.post_artifact(
        task_id=task_id,
        artifact_name="Layout de Reapresentação.xlsx",
        filepath=path
    )

if __name__ == '__main__':
    main()