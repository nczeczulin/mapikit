#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <frameobject.h>


typedef struct {
    PyObject_HEAD
    PyObject *func;
    PyObject *result_handler;
    PyObject *error_handler;
} CallWrapperObject;

static PyTypeObject CallWrapper_Type;

static void
CallWrapper_dealloc(CallWrapperObject *self)
{
    Py_CLEAR(self->func);
    Py_CLEAR(self->result_handler);
    Py_CLEAR(self->error_handler);
    Py_TYPE(self)->tp_free((PyObject *) self);
}

static PyObject *
CallWrapper_new(PyTypeObject *type, PyObject *args, PyObject *kwargs)
{
    CallWrapperObject *self;
    self = (CallWrapperObject *) type->tp_alloc(type, 0);
    if (self != NULL) {
        self->func = NULL;
        self->result_handler = NULL;
        self->error_handler = NULL;
    }
    return (PyObject *) self;
}

PyDoc_STRVAR(CallWrapper_doc,
    "CallWrapper\n"
    "\n"
    "Args:\n"
    "    func (callable): The function being called.\n"
    "    result_handler (callable): The function to process returned results from func.\n"
    "    error_handler (callable, optional): Called passing the exc value if func raises an error.\n"
    "Returns:\n"
    "    Depends on the result_handler being used.");

static int
CallWrapper_init(CallWrapperObject *self, PyObject *args, PyObject *kwargs)
{
    static char *kwlist[] = {"func", "result_handler", "error_handler", NULL};
    static const char *type_err_format = "'%s' object is not callable";
    PyObject *func = NULL, *result_handler = NULL, *error_handler = NULL, *tmp;

    if (!PyArg_ParseTupleAndKeywords(args, kwargs, "OO|O:CallWrapper", kwlist,
                                     &func, &result_handler, &error_handler))
        return -1;

    if (!PyCallable_Check(func)) {
        PyErr_Format(PyExc_TypeError, type_err_format, Py_TYPE(func)->tp_name);
        return -1;
    }
    if (!PyCallable_Check(result_handler)) {
        PyErr_Format(PyExc_TypeError, type_err_format, Py_TYPE(result_handler)->tp_name);
        return -1;
    }
    if (!PyCallable_Check(error_handler)) {
        PyErr_Format(PyExc_TypeError, type_err_format, Py_TYPE(error_handler)->tp_name);
        return -1;
    }

    tmp = self->func;
    Py_INCREF(func);
    self->func = func;
    Py_XDECREF(tmp);

    tmp = self->result_handler;
    Py_INCREF(result_handler);
    self->result_handler = result_handler;
    Py_XDECREF(tmp);

    if (error_handler) {
        tmp = self->error_handler;
        Py_INCREF(error_handler);
        self->error_handler = error_handler;
        Py_XDECREF(tmp);
    }
    return 0;
}

PyDoc_STRVAR(CallWrapper_wrapper_doc,
    "wrapper\n"
    "\n"
    "Args:\n"
    "    *args (optional): Passed to func\n"
    "    **kwargs (optional): Passed to func\n"
    "Returns:\n"
    "    Depends on the result_handler being used.\n"
    "Raises:\n"
    "    Depends on the func and handlers being used.");

static PyObject *
CallWrapper_call(CallWrapperObject *self, PyObject *args, PyObject *kwargs)
{
    PyObject *result, *handler_args;

    if (!(handler_args = PyTuple_New(1)))
        return NULL;

    if (!(result = PyObject_Call(self->func, args, kwargs))) {
        if (self->error_handler) {
            PyObject *exc_type = NULL, *exc_val = NULL, *exc_tb = NULL;

            PyErr_Fetch(&exc_type, &exc_val, &exc_tb);
            PyErr_NormalizeException(&exc_type, &exc_val, &exc_tb);
            if (exc_tb != NULL) {
                PyException_SetTraceback(exc_val, exc_tb);
            }

            Py_INCREF(exc_val);
            PyTuple_SET_ITEM(handler_args, 0, exc_val);
            result = PyObject_CallObject(self->error_handler, handler_args);
            Py_CLEAR(result);
            _PyErr_ChainExceptions(exc_type, exc_val, exc_tb);
        }
    }
    else {
        PyTuple_SET_ITEM(handler_args, 0, result);
        result = PyObject_CallObject(self->result_handler, handler_args);
    }

    Py_CLEAR(handler_args);
    return result;
}

static PyMethodDef CallWrapper_methods[] = {
    {NULL}
};

static PyTypeObject CallWrapper_Type = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name = "mapikit.callwrapper.CallWrapper",
    .tp_doc = CallWrapper_doc,
    .tp_basicsize = sizeof(CallWrapperObject),
    .tp_itemsize = 0,
    .tp_flags = Py_TPFLAGS_DEFAULT | Py_TPFLAGS_BASETYPE,
    .tp_new = CallWrapper_new,
    .tp_init = (initproc) CallWrapper_init,
    .tp_call = (ternaryfunc) CallWrapper_call,
    .tp_dealloc = (destructor) CallWrapper_dealloc,
    .tp_methods = CallWrapper_methods,
};

static int
callwrapper_exec(PyObject *m)
{
    if (PyType_Ready(&CallWrapper_Type) < 0)
        goto fail;
    PyModule_AddObject(m, "CallWrapper", (PyObject *)&CallWrapper_Type);

    return 0;
 fail:
    Py_XDECREF(m);
    return -1;
}

static struct PyModuleDef_Slot callwrapper_slots[] = {
    {Py_mod_exec, callwrapper_exec},
    {0, NULL},
};

PyDoc_STRVAR(callwrappermodule_doc,
    "callwrapper c-extension module.\n"
    "\n"
    "...");

static PyModuleDef callwrappermodule_def = {
    PyModuleDef_HEAD_INIT,  /* m_base */
    "mapikit.callwrapper", /* m_name */
    callwrappermodule_doc, /* m_doc */
    0,    /* m_size */
    NULL, /* m_methods */
    callwrapper_slots, /* m_slots */
    NULL, /* m_traverse */
    NULL, /* m_clear */
    NULL, /* m_free */
};

PyMODINIT_FUNC
PyInit_callwrapper(void)
{
    return PyModuleDef_Init(&callwrappermodule_def);
}
