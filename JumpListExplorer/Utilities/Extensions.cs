using System;
using System.Runtime.InteropServices.ComTypes;
using JumpListExplorer.Interop;

namespace JumpListExplorer.Utilities
{
    internal static class Extensions
    {
        public static HRESULT AddToBindCtx(this IBindCtx bindContext, string name, object value, bool throwOnError = true)
        {
            ArgumentNullException.ThrowIfNull(bindContext);
            ArgumentNullException.ThrowIfNull(name);
            var hr = bindContext.EnsureBindCtxPropertyBag(out var ps, throwOnError);
            if (hr != 0)
                return hr;

            if (ps is not IPropertyBag bag)
                return HRESULTS.E_FAIL;

            var variant = value;
            return bag.Write(name, ref variant).ThrowOnError(throwOnError);
        }

        public static HRESULT EnsureBindCtxPropertyBag(this IBindCtx bindContext, out object? store, bool throwOnError = true)
        {
            ArgumentNullException.ThrowIfNull(bindContext);

            store = null;
            const string STR_PROPERTYBAG_PARAM = "SHBindCtxPropertyBag";
            object? unk = null;
            try
            {
                bindContext.GetObjectParam(STR_PROPERTYBAG_PARAM, out unk);
            }
            catch
            {
                // continue
            }

            if (unk == null)
            {
                var hr = Native.PSCreateMemoryPropertyStore(typeof(IPropertyStore).GUID, out var ps).ThrowOnError(throwOnError);
                if (hr != 0)
                    return hr;

                store = ps;
                if (throwOnError)
                {
                    bindContext.RegisterObjectParam(STR_PROPERTYBAG_PARAM, store);
                }
                else
                {
                    try
                    {
                        bindContext.RegisterObjectParam(STR_PROPERTYBAG_PARAM, store);
                    }
                    catch
                    {
                        return HRESULTS.E_FAIL;
                    }
                }
            }
            return HRESULTS.S_OK;
        }
    }
}
