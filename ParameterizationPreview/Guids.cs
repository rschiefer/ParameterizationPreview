// Guids.cs
// MUST match guids.h
using System;

namespace Company.ParameterizationPreview
{
    static class GuidList
    {
        public const string guidParameterizationPreviewPkgString = "924827a5-4d34-44bf-a250-aa72d2b892b6";
        public const string guidParameterizationPreviewCmdSetString = "12d5c20a-c12b-4838-8290-b7f98a3ce785";

        public static readonly Guid guidParameterizationPreviewCmdSet = new Guid(guidParameterizationPreviewCmdSetString);
    };
}