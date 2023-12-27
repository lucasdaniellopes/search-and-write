import time

class APIUtils:
    @staticmethod
    def retry_api_call(api_function, *args, **kwargs):
        last_request = None  # Reinicializa a variável
        retry_delay_seconds = 5  # Inicializa o valor de espera

        while APIUtils.requests_count < 60:
            elapsed_time = time.time() - APIUtils.start_time

            if elapsed_time >= 60:
                APIUtils.requests_count = 0
                APIUtils.start_time = time.time()

            try:
                result = api_function(*args, **kwargs)
                APIUtils.requests_count += 1
                print("QUASE LAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")
                return result
            except Exception as e:
                if 'RATE_LIMIT_EXCEEDED' in str(e) or (hasattr(e, 'code') and e.code == 429):
                    print("ENTROOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOU")
                    print(f'Erro: {str(e)}. Aguardando...')

                    # Aplica espera exponencial antes de reenviar a solicitação
                    time.sleep(retry_delay_seconds)

                    # Aumenta o tempo de espera exponencial para a próxima tentativa
                    retry_delay_seconds *= 2
                    continue
                else:
                    print(f'Erro: {str(e)}. Tentando novamente...')
                    last_request = {'api_function': api_function, 'args': args, 'kwargs': kwargs}
                    retry_delay_seconds *= 2
                    continue

        # Se chegou a este ponto, todas as tentativas foram esgotadas
        # Tentar novamente a última chamada que falhou
        if last_request:
            print('Tentando novamente a última chamada que falhou...')
            APIUtils.retry_api_call(last_request['api_function'], *last_request['args'], **last_request['kwargs'])
        else:
            print("Limite de tentativas atingido. Não foi possível realizar a chamada.")
