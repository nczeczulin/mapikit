#include <windows.h>
#define PY_SSIZE_T_CLEAN
#include <Python.h>

/* MAPIDefS.h */
#define PROP_TYPE_MASK                              ((ULONG)0x0000FFFF)
#define PROP_TYPE(ulPropTag)                        (((ULONG)(ulPropTag))&PROP_TYPE_MASK)
#define PROP_ID(ulPropTag)                          (((ULONG)(ulPropTag))>>16)
#define PROP_TAG(ulPropType,ulPropID)               ((((ULONG)(ulPropID))<<16)|((ULONG)(ulPropType)))
#define CHANGE_PROP_TYPE(ulPropTag, ulPropType)     (((ULONG)0xFFFF0000 & (ulPropTag)) | (ulPropType))

PyDoc_STRVAR(macros_PROP_TYPE_doc,
    "PROP_TYPE(proptag)\n"
    "--\n\n"
    "Returns the property type of a specified property tag.\n"
    "\n"
    "Args:\n"
    "    proptag (int): Property tag that contains the property type to be returned.\n"
    "Returns:\n"
    "    int: The property type contained in the property tag.");

static PyObject *
macros_PROP_TYPE(PyObject *self, PyObject *arg)
{
    ULONG ulPropTag = PyLong_AsUnsignedLongMask(arg);

    if (ulPropTag == -1 && PyErr_Occurred())
        return NULL;
    return PyLong_FromUnsignedLong(PROP_TYPE(ulPropTag));
}

PyDoc_STRVAR(macros_PROP_ID_doc,
    "PROP_ID(proptag)\n"
    "--\n\n"
    "Returns the property identifier of a specified property tag.\n"
    "\n"
    "Args:\n"
    "    proptag (int): Property tag that contains the property identifier to be returned.\n"
    "Returns:\n"
    "    int: The property identifier contained in the property tag.");

static PyObject *
macros_PROP_ID(PyObject *self, PyObject *arg)
{
    ULONG ulPropTag = PyLong_AsUnsignedLongMask(arg);

    if (ulPropTag == -1 && PyErr_Occurred())
        return NULL;
    return PyLong_FromUnsignedLong(PROP_ID(ulPropTag));
}

PyDoc_STRVAR(macros_PROP_TYPE_AND_ID_doc,
    "PROP_TYPE_AND_ID(proptag)\n"
    "--\n\n"
    "Returns the property type and identifier of a specified property tag.\n"
    "\n"
    "Args:\n"
    "    proptag (int): Property tag that contains the property identifier to be returned.\n"
    "Returns:\n"
    "    (int, int): The property type and identifier contained in the property tag.");

static PyObject *
macros_PROP_TYPE_AND_ID(PyObject *self, PyObject *arg)
{
    ULONG ulPropTag = PyLong_AsUnsignedLongMask(arg);
    PyObject *result;

    if (ulPropTag == -1 && PyErr_Occurred())
        return NULL;

    if (!(result = PyTuple_New(2)))
        return NULL;

    PyTuple_SET_ITEM(result, 0, PyLong_FromUnsignedLong(PROP_TYPE(ulPropTag)));
    PyTuple_SET_ITEM(result, 1, PyLong_FromUnsignedLong(PROP_ID(ulPropTag)));

    return result;
}

PyDoc_STRVAR(macros_PROP_TAG_doc,
    "PROP_TAG(proptype, propid)\n"
    "--\n\n"
    "Returns a property tag created by combining a specified property type and identifier.\n"
    "\n"
    "Args:\n"
    "    proptype (int): Property type for the new property tag.\n"
    "    propid (int): Property identifier for the new property tag.\n"
    "Returns:\n"
    "    int: The property tag created by combining the property type and identifier.");

static PyObject *
macros_PROP_TAG(PyObject *self, PyObject *args)
{
    PyObject *obPropType;
    PyObject *obPropID;

    if (PyArg_UnpackTuple(args, "PROP_TAG", 2, 2, &obPropType, &obPropID))
    {
        ULONG ulPropType = PyLong_AsUnsignedLongMask(obPropType);
        if (ulPropType == -1 && PyErr_Occurred())
            return NULL;
        ULONG ulPropID = PyLong_AsUnsignedLongMask(obPropID);
        if (ulPropID == -1 && PyErr_Occurred())
            return NULL;
        return PyLong_FromUnsignedLong(PROP_TAG(ulPropType, ulPropID));
    }
    return NULL;
}

PyDoc_STRVAR(macros_CHANGE_PROP_TYPE_doc,
    "CHANGE_PROP_TYPE(proptag, proptype)\n"
    "--\n\n"
    "Updates the property type of a specified property tag.\n"
    "\n"
    "Args:\n"
    "    proptag (int): The property tag to be modified.\n"
    "    proptype (int): The new value for the property type.\n"
    "Returns:\n"
    "    int: The modified property tag with an updated property type.");

static PyObject *
macros_CHANGE_PROP_TYPE(PyObject *self, PyObject *args)
{
    PyObject *obPropTag;
    PyObject *obPropType;

    if (PyArg_UnpackTuple(args, "CHANGE_PROP_TYPE", 2, 2, &obPropTag, &obPropType))
    {
        ULONG ulPropTag = PyLong_AsUnsignedLongMask(obPropTag);
        if (ulPropTag == -1 && PyErr_Occurred())
            return NULL;
        ULONG ulPropType = PyLong_AsUnsignedLongMask(obPropType);
        if (ulPropType == -1 && PyErr_Occurred())
            return NULL;
        return PyLong_FromUnsignedLong(CHANGE_PROP_TYPE(ulPropTag, ulPropType));
    }
    return NULL;
}

static PyMethodDef macros_methods[] = {
    { "PROP_TYPE", (PyCFunction)macros_PROP_TYPE, METH_O, macros_PROP_TYPE_doc},
    { "PROP_ID", (PyCFunction)macros_PROP_ID, METH_O, macros_PROP_ID_doc},
    { "PROP_TYPE_AND_ID", (PyCFunction)macros_PROP_TYPE_AND_ID, METH_O, macros_PROP_TYPE_AND_ID_doc},
    { "PROP_TAG", (PyCFunction)macros_PROP_TAG, METH_VARARGS, macros_PROP_TAG_doc},
    { "CHANGE_PROP_TYPE", (PyCFunction)macros_CHANGE_PROP_TYPE, METH_VARARGS, macros_CHANGE_PROP_TYPE_doc},
    { NULL, NULL, 0, NULL }
};

PyDoc_STRVAR(macros_doc,
    "MAPI macros implemented in a c-extension module.\n"
    "\n"
    "They are a little faster than the ones in the Pywin32 mapitags module\n"
    "that are written in Python.");

static PyModuleDef macros_def = {
    PyModuleDef_HEAD_INIT,
    "mapikit.macros",
    macros_doc,
    0,    /* m_size */
    macros_methods, /* m_methods */
    NULL, /* m_slots */
    NULL, /* m_traverse */
    NULL, /* m_clear */
    NULL, /* m_free */
};

PyMODINIT_FUNC
PyInit_macros()
{
    return PyModuleDef_Init(&macros_def);
}
