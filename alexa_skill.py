import logging
import os

from ask_sdk_core.skill_builder import SkillBuilder
from ask_sdk_core.dispatch_components import AbstractRequestHandler, AbstractExceptionHandler
from ask_sdk_core.utils import is_request_type, is_intent_name
from ask_sdk_model import Response
from ask_sdk_core.handler_input import HandlerInput

import google.generativeai as genai

# Configura il logging per il debug
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Configurazione della chiave API per Google AI Studio
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY")
if GOOGLE_API_KEY:
    genai.configure(api_key=GOOGLE_API_KEY)
else:
    logger.error("La variabile d'ambiente GOOGLE_API_KEY non è impostata.")

# Configurazione del modello Gemini
MODEL_NAME = "gemini-2.0-flash"
generation_config = genai.types.GenerationConfig(
    temperature=0.7,
    top_p=0.8,
    top_k=40,
    max_output_tokens=256,
)
safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# Handler per l'intent "InterrogaGeminiIntent"
class InterrogaGeminiIntentHandler(AbstractRequestHandler):
    def can_handle(self, handler_input: HandlerInput) -> bool:
        return is_intent_name("InterrogaGeminiIntent")(handler_input)

    def handle(self, handler_input: HandlerInput) -> Response:
        slots = handler_input.request_envelope.request.intent.slots
        prompt = slots.get("prompt").value if slots and "prompt" in slots else None

        if prompt:
            try:
                model = genai.GenerativeModel(
                    MODEL_NAME,
                    generation_config=generation_config,
                    safety_settings=safety_settings,
                )
                response = model.generate_content(prompt)
                gemini_response = response.text

                speech_text = gemini_response
                reprompt_text = "Posso aiutarti con qualcos'altro?"
                return (
                    handler_input.response_builder
                    .speak(speech_text)
                    .ask(reprompt_text)
                    .response
                )
            except Exception as e:
                logger.error(f"Errore durante la chiamata all'API di Gemini: {e}")
                speech_text = "Si è verificato un errore durante l'interrogazione di Gemini. Riprova più tardi."
                return handler_input.response_builder.speak(speech_text).response
        else:
            speech_text = "Non ho capito la tua domanda. Riprova."
            return handler_input.response_builder.speak(speech_text).response

# Handler per la richiesta di lancio della skill
class LaunchRequestHandler(AbstractRequestHandler):
    def can_handle(self, handler_input: HandlerInput) -> bool:
        return is_request_type("LaunchRequest")(handler_input)

    def handle(self, handler_input: HandlerInput) -> Response:
        speech_text = "Benvenuto! Puoi chiedermi qualsiasi cosa."
        reprompt_text = "Cosa vorresti chiedere?"
        return (
            handler_input.response_builder
            .speak(speech_text)
            .ask(reprompt_text)
            .response
        )

# Handler per la richiesta di aiuto
class HelpIntentHandler(AbstractRequestHandler):
    def can_handle(self, handler_input: HandlerInput) -> bool:
        return is_intent_name("AMAZON.HelpIntent")(handler_input)

    def handle(self, handler_input: HandlerInput) -> Response:
        speech_text = "Puoi chiedermi qualsiasi cosa e io cercherò di risponderti usando Gemini."
        reprompt_text = "Cosa vorresti chiedere?"
        return (
            handler_input.response_builder
            .speak(speech_text)
            .ask(reprompt_text)
            .response
        )

# Handler per la richiesta di annullamento o stop
class CancelOrStopIntentHandler(AbstractRequestHandler):
    def can_handle(self, handler_input: HandlerInput) -> bool:
        return (
            is_intent_name("AMAZON.CancelIntent")(handler_input)
            or is_intent_name("AMAZON.StopIntent")(handler_input)
        )

    def handle(self, handler_input: HandlerInput) -> Response:
        speech_text = "A presto!"
        return handler_input.response_builder.speak(speech_text).response

# Handler per la richiesta di sessione terminata
class SessionEndedRequestHandler(AbstractRequestHandler):
    def can_handle(self, handler_input: HandlerInput) -> bool:
        return is_request_type("SessionEndedRequest")(handler_input)

    def handle(self, handler_input: HandlerInput) -> Response:
        logger.info(f"Session ended with reason: {handler_input.request_envelope.request.reason}")
        return handler_input.response_builder.response

# Exception handler per catturare tutte le eccezioni
class CatchAllExceptionHandler(AbstractExceptionHandler):
    def can_handle(self, handler_input, exception) -> bool:
        return True

    def handle(self, handler_input, exception):
        logger.error(f"Errore imprevisto: {exception}", exc_info=True)
        speech_text = "Si è verificato un errore imprevisto. Riprova più tardi."
        return handler_input.response_builder.speak(speech_text).response

# Costruzione della skill
sb = SkillBuilder()
sb.add_request_handler(LaunchRequestHandler())
sb.add_request_handler(HelpIntentHandler())
sb.add_request_handler(CancelOrStopIntentHandler())
sb.add_request_handler(SessionEndedRequestHandler())
sb.add_request_handler(InterrogaGeminiIntentHandler())
sb.add_exception_handler(CatchAllExceptionHandler())

handler = sb.lambda_handler()
