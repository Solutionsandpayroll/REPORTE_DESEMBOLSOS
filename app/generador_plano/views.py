import json
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from .utils import (
    leer_formato_envio, leer_maestro_externo, aplicar_manuales,
    generar_santander, generar_bancolombia,
    BANCOS_SANTANDER, BANCOS_BANCOLOMBIA,
    TIPO_DOC, APLICACIONES, TIPOS_PAGO, TIPOS_CTA_DEB,
    TIPOS_DOC_BCOL, TIPOS_TRANSACCION, HDR_DEFAULT,
    PAB_NOM2COD,
)


def index(request):
    return render(request, 'index.html')


@csrf_exempt
def preview_datos(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'Método no permitido'}, status=405)
    archivo = request.FILES.get('formato_envio')
    if not archivo:
        return JsonResponse({'error': 'No se recibió el archivo'}, status=400)

    maestro = None
    if request.FILES.get('maestro_empleados'):
        try:
            maestro = leer_maestro_externo(request.FILES['maestro_empleados'])
        except Exception as e:
            return JsonResponse({'error': f'Error maestro: {e}'}, status=400)

    try:
        datos = leer_formato_envio(archivo, maestro_externo=maestro)
        return JsonResponse({
            'ok':                True,
            'fecha':             datos['fecha'],
            'consecutivo':       datos['consecutivo'],
            'cantidad':          datos['cantidad'],
            'total':             datos['total'],
            'registros':         datos['registros'],
            'faltantes':         datos['faltantes'],
            'cat_bancos_sant':   BANCOS_SANTANDER,
            'cat_bancos_bcol':   BANCOS_BANCOLOMBIA,
            'cat_tdoc_keys':     list(TIPO_DOC.keys()),
            'cat_tdoc_labels':   [v[0] for v in TIPO_DOC.values()],
            'cat_aplicaciones':  APLICACIONES,
            'cat_tipos_pago':    TIPOS_PAGO,
            'cat_tipos_cta_deb': TIPOS_CTA_DEB,
            'cat_tipos_doc_bcol': TIPOS_DOC_BCOL,
            'cat_tipos_transac': TIPOS_TRANSACCION,
            'hdr_default':       HDR_DEFAULT,
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@csrf_exempt
def generar_archivo(request):
    if request.method != 'POST':
        return HttpResponse('Método no permitido', status=405)
    archivo = request.FILES.get('formato_envio')
    if not archivo:
        return HttpResponse('No se recibió el archivo', status=400)

    plantilla     = request.POST.get('plantilla', 'santander')
    referencia    = request.POST.get('referencia', '').strip()
    manuales_json = request.POST.get('datos_manuales', '{}') or '{}'
    hdr_json      = request.POST.get('header_config',  '{}') or '{}'

    maestro = None
    if request.FILES.get('maestro_empleados'):
        try:
            maestro = leer_maestro_externo(request.FILES['maestro_empleados'])
        except Exception as e:
            return HttpResponse(f'Error maestro: {e}', status=400)

    try:
        manuales   = json.loads(manuales_json)
        hdr_config = json.loads(hdr_json)
    except Exception:
        manuales = {}
        hdr_config = {}

    try:
        datos = leer_formato_envio(archivo, maestro_externo=maestro)
        if manuales:
            datos = aplicar_manuales(datos, manuales)

        consecutivo = datos.get('consecutivo', '').replace('/', '-').replace(' ', '_')

        if plantilla == 'bancolombia':
            hdr = HDR_DEFAULT.copy()
            hdr.update(hdr_config)
            if referencia:
                hdr['descripcion'] = referencia
            output   = generar_bancolombia(datos, hdr)
            filename = f'Bancolombia_PAB_{consecutivo}.xlsx' if consecutivo else 'Bancolombia_PAB.xlsx'
            ctype    = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            # ← CORRECTO: el argumento se llama 'referencia', no 'referencia_base'
            output   = generar_santander(datos, referencia)
            filename = f'Santander_{consecutivo}.xlsm' if consecutivo else 'Santander.xlsm'
            ctype    = 'application/vnd.ms-excel.sheet.macroEnabled.12'

        resp = HttpResponse(output.read(), content_type=ctype)
        resp['Content-Disposition'] = f'attachment; filename="{filename}"'
        return resp

    except Exception as e:
        return HttpResponse(f'Error: {e}', status=500)
